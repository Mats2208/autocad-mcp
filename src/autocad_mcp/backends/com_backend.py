"""COM (ActiveX) backend for full AutoCAD on Windows — local addition.

Draws through AutoCAD's official ActiveX object model instead of typing LISP
commands into the command line. Compared to the File IPC backend this:

- does not depend on keyboard focus (the user can keep using the mouse),
- is immune to running object snaps (OSNAP) and to the UI language,
- returns AutoCAD's real error messages instead of empty strings,
- supports block definitions (``block_define``), which File IPC lacks.

Operations with no ActiveX equivalent (fillet, chamfer, purge, plot, P&ID,
free LISP, undo/redo, screenshots) are inherited from ``FileIPCBackend``, so the
LISP dispatcher must still be loaded; ``initialize`` loads it if it is missing.
"""

from __future__ import annotations

import math
import time
from typing import Any

import structlog

from autocad_mcp.backends.base import CommandResult
from autocad_mcp.backends.file_ipc import FileIPCBackend
from autocad_mcp.config import LISP_DIR

log = structlog.get_logger()

# ACI color names accepted by the LISP backend, mapped to indices.
_COLOR_NAMES = {
    "red": 1, "yellow": 2, "green": 3, "cyan": 4, "blue": 5,
    "magenta": 6, "white": 7, "black": 7, "gray": 8, "grey": 8,
    "lightgray": 9, "lightgrey": 9, "bylayer": 256, "byblock": 0,
}

# AcLineWeight values are hundredths of a millimetre (e.g. 0.50 mm -> 50).
_LINEWEIGHTS = {0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211}

_RPC_RETRY = (-2147418111, -2147417846)  # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER


class ComBackend(FileIPCBackend):
    """Full AutoCAD via ActiveX, with File IPC as fallback for the rest."""

    def __init__(self):
        super().__init__()
        self._app = None
        self._last_handle: str | None = None

    @property
    def name(self) -> str:
        return "com"

    # ------------------------------------------------------------------
    # COM plumbing
    # ------------------------------------------------------------------

    def _connect(self):
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        if self._app is None:
            self._app = win32com.client.GetActiveObject("AutoCAD.Application")
        return self._app

    @property
    def _doc(self):
        return self._connect().ActiveDocument

    @property
    def _ms(self):
        return self._doc.ModelSpace

    @staticmethod
    def _pt(x: float, y: float, z: float = 0.0):
        import pythoncom
        import win32com.client

        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))

    @staticmethod
    def _darr(values):
        import pythoncom
        import win32com.client

        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, tuple(float(v) for v in values))

    @staticmethod
    def _objarr(objs):
        import pythoncom
        import win32com.client

        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, tuple(objs))

    @staticmethod
    def _err(e: Exception) -> str:
        """Extract AutoCAD's message from a COM error."""
        try:
            info = e.excepinfo  # type: ignore[attr-defined]
            if info and info[2]:
                return str(info[2]).strip()
            return f"{e.args[1]} (hresult {e.args[0]})"  # type: ignore[index]
        except Exception:
            return str(e) or e.__class__.__name__

    async def _run(self, fn, *, undo: bool = True) -> CommandResult:
        """Run a COM operation under the backend lock, retrying while AutoCAD is busy.

        Each modifying operation is wrapped in an undo mark so ``drawing(undo)``
        reverts exactly one MCP call.
        """
        import pythoncom

        async with self._lock:
            for attempt in range(40):
                try:
                    self._connect()
                    doc = self._doc
                    if undo:
                        doc.StartUndoMark()
                    try:
                        payload = fn()
                    finally:
                        if undo:
                            doc.EndUndoMark()
                    return CommandResult(ok=True, payload=payload)
                except pythoncom.com_error as e:
                    if e.hresult in _RPC_RETRY and attempt < 39:
                        time.sleep(0.25)  # AutoCAD busy (modal dialog, active command)
                        continue
                    if e.hresult in (-2147221021, -2147023174, -2147023170):
                        self._app = None  # AutoCAD restarted or COM server gone
                    return CommandResult(ok=False, error=self._err(e))
                except Exception as e:
                    return CommandResult(ok=False, error=self._err(e))
        return CommandResult(ok=False, error="AutoCAD stayed busy; try again")

    def _obj(self, entity_id: str):
        if entity_id == "last":
            if not self._last_handle:
                raise ValueError("No entity created yet in this session ('last')")
            entity_id = self._last_handle
        try:
            return self._doc.HandleToObject(entity_id)
        except Exception:
            raise ValueError(f"Entity '{entity_id}' not found")

    def _ensure_layer(self, name: str | None):
        if not name:
            return
        layers = self._doc.Layers
        try:
            layers.Item(name)
        except Exception:
            layers.Add(name)

    def _ensure_linetype(self, name: str):
        if not name or name.upper() in ("CONTINUOUS", "BYLAYER", "BYBLOCK"):
            return
        lts = self._doc.Linetypes
        try:
            lts.Item(name)
            return
        except Exception:
            pass
        last_error = None
        for lin in ("acadiso.lin", "acad.lin"):
            try:
                lts.Load(name, lin)
                return
            except Exception as e:
                last_error = e
        raise ValueError(f"Linetype '{name}' not found in acadiso.lin/acad.lin: {self._err(last_error)}")

    @staticmethod
    def _color(color: str | int | None) -> int | None:
        if color is None or color == "":
            return None
        if isinstance(color, int):
            return color
        c = str(color).strip().lower()
        if c.isdigit():
            return int(c)
        if c in _COLOR_NAMES:
            return _COLOR_NAMES[c]
        raise ValueError(f"Unknown color '{color}' (use an ACI number 0-256 or a basic color name)")

    def _done(self, obj, layer: str | None = None, extra: dict | None = None) -> dict:
        if layer:
            self._ensure_layer(layer)
            obj.Layer = layer
        h = obj.Handle
        self._last_handle = h
        out = {"entity_type": self._dxf_type(obj), "handle": h}
        if extra:
            out.update(extra)
        return out

    @staticmethod
    def _dxf_type(obj) -> str:
        name = obj.ObjectName  # e.g. AcDbLine
        return {
            "AcDbLine": "LINE", "AcDbCircle": "CIRCLE", "AcDbArc": "ARC",
            "AcDbPolyline": "LWPOLYLINE", "AcDbEllipse": "ELLIPSE", "AcDbText": "TEXT",
            "AcDbMText": "MTEXT", "AcDbHatch": "HATCH", "AcDbBlockReference": "INSERT",
            "AcDbRotatedDimension": "DIMENSION", "AcDbAlignedDimension": "DIMENSION",
            "AcDb2LineAngularDimension": "DIMENSION", "AcDb3PointAngularDimension": "DIMENSION",
            "AcDbRadialDimension": "DIMENSION", "AcDbLeader": "LEADER",
        }.get(name, name.replace("AcDb", "").upper())

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    async def initialize(self) -> CommandResult:
        try:
            self._connect()
            _ = self._doc.Name
        except Exception as e:
            return CommandResult(ok=False, error=f"AutoCAD COM not available: {self._err(e)}")
        result = await super().initialize()
        if not result.ok:
            # The LISP dispatcher (needed for the inherited operations) is not
            # loaded in this drawing: load it through COM and retry once.
            lisp = str(LISP_DIR / "mcp_dispatch.lsp").replace("\\", "/")
            try:
                self._doc.SendCommand(f'(load "{lisp}")\n')
                time.sleep(1.0)
            except Exception as e:
                return CommandResult(ok=False, error=f"Could not load {lisp}: {self._err(e)}")
            result = await super().initialize()
        return result

    async def status(self) -> CommandResult:
        result = await super().status()
        if result.ok and isinstance(result.payload, dict):
            result.payload["backend"] = self.name
            try:
                result.payload["drawing"] = self._doc.Name
                result.payload["autocad_version"] = self._connect().Version
            except Exception:
                pass
        return result

    # ------------------------------------------------------------------
    # Drawing management
    # ------------------------------------------------------------------

    async def drawing_info(self) -> CommandResult:
        def fn():
            doc = self._doc
            ms = doc.ModelSpace
            try:
                mn, mx = doc.GetVariable("EXTMIN"), doc.GetVariable("EXTMAX")
                extents = {"min": list(mn)[:2], "max": list(mx)[:2]}
            except Exception:
                extents = None
            return {
                "drawing": doc.Name,
                "path": doc.FullName,
                "entity_count": ms.Count,
                "layers": [doc.Layers.Item(i).Name for i in range(doc.Layers.Count)],
                "blocks": self._user_blocks(),
                "extents": extents,
                "units": {"INSUNITS": doc.GetVariable("INSUNITS")},
            }
        return await self._run(fn, undo=False)

    async def drawing_save(self, path: str | None = None) -> CommandResult:
        def fn():
            doc = self._doc
            if path:
                p = path.replace("/", "\\")
                if not p.lower().endswith(".dwg"):
                    p += ".dwg"
                doc.SaveAs(p)
            else:
                doc.Save()
            return {"saved": doc.FullName}
        return await self._run(fn, undo=False)

    async def drawing_get_variables(self, names: list[str] | None = None) -> CommandResult:
        names = [n.lstrip("$") for n in (names or ["ACADVER", "INSUNITS", "DWGNAME", "CLAYER", "LTSCALE", "DIMSCALE"])]

        def fn():
            out: dict[str, Any] = {}
            for n in names:
                try:
                    v = self._doc.GetVariable(n)
                    out[n] = list(v) if isinstance(v, tuple) else v
                except Exception:
                    out[n] = None
            return out
        return await self._run(fn, undo=False)

    async def drawing_open(self, path: str) -> CommandResult:
        def fn():
            doc = self._connect().Documents.Open(path.replace("/", "\\"))
            return {"drawing": doc.Name, "path": doc.FullName}
        result = await self._run(fn, undo=False)
        if result.ok:
            # Each document has its own LISP namespace: make sure the dispatcher is there.
            ping = await self._dispatch("ping", {})
            if not ping.ok:
                lisp = str(LISP_DIR / "mcp_dispatch.lsp").replace("\\", "/")
                self._doc.SendCommand(f'(load "{lisp}")\n')
        return result

    # ------------------------------------------------------------------
    # Entity creation
    # ------------------------------------------------------------------

    async def create_line(self, x1, y1, x2, y2, layer=None) -> CommandResult:
        return await self._run(lambda: self._done(self._ms.AddLine(self._pt(x1, y1), self._pt(x2, y2)), layer))

    async def create_circle(self, cx, cy, radius, layer=None) -> CommandResult:
        return await self._run(lambda: self._done(self._ms.AddCircle(self._pt(cx, cy), float(radius)), layer))

    async def create_polyline(self, points, closed=False, layer=None) -> CommandResult:
        def fn():
            if len(points) < 2:
                raise ValueError("A polyline needs at least 2 points")
            flat = [c for p in points for c in (p[0], p[1])]
            pl = self._ms.AddLightWeightPolyline(self._darr(flat))
            pl.Closed = bool(closed)
            return self._done(pl, layer)
        return await self._run(fn)

    async def create_rectangle(self, x1, y1, x2, y2, layer=None) -> CommandResult:
        return await self.create_polyline([[x1, y1], [x2, y1], [x2, y2], [x1, y2]], closed=True, layer=layer)

    async def create_arc(self, cx, cy, radius, start_angle, end_angle, layer=None) -> CommandResult:
        return await self._run(lambda: self._done(self._ms.AddArc(
            self._pt(cx, cy), float(radius), math.radians(start_angle), math.radians(end_angle)), layer))

    async def create_ellipse(self, cx, cy, major_x, major_y, ratio, layer=None) -> CommandResult:
        # Same convention as the LISP backend: (major_x, major_y) is the end point of the major axis.
        return await self._run(lambda: self._done(self._ms.AddEllipse(
            self._pt(cx, cy), self._pt(major_x - cx, major_y - cy), float(ratio)), layer))

    async def create_mtext(self, x, y, width, text, height=2.5, layer=None) -> CommandResult:
        def fn():
            mt = self._ms.AddMText(self._pt(x, y), float(width), str(text))
            mt.Height = float(height)
            return self._done(mt, layer)
        return await self._run(fn)

    async def create_hatch(self, entity_id, pattern="ANSI31") -> CommandResult:
        def fn():
            boundary = self._obj(entity_id)
            pat = (pattern or "ANSI31").upper()
            # acHatchPatternTypePreDefined = 1; associative so it follows the boundary.
            h = self._ms.AddHatch(1, pat, True)
            h.AppendOuterLoop(self._objarr([boundary]))
            if pat != "SOLID":
                h.PatternScale = float(self._doc.GetVariable("HPSCALE") or 1.0)
            h.Evaluate()
            h.Layer = self._doc.ActiveLayer.Name
            return self._done(h)
        return await self._run(fn)

    # ------------------------------------------------------------------
    # Entity queries and modification
    # ------------------------------------------------------------------

    def _describe(self, obj) -> dict:
        d: dict[str, Any] = {"handle": obj.Handle, "type": self._dxf_type(obj), "layer": obj.Layer}
        for prop, key in (("StartPoint", "start"), ("EndPoint", "end"), ("Center", "center"),
                          ("InsertionPoint", "insertion"), ("Radius", "radius"), ("TextString", "text"),
                          ("Name", "block"), ("Area", "area"), ("Length", "length"), ("Closed", "closed"),
                          ("Measurement", "measurement"), ("Rotation", "rotation")):
            try:
                v = getattr(obj, prop)
                d[key] = [round(c, 6) for c in v][:2] if isinstance(v, tuple) else v
            except Exception:
                pass
        if d["type"] == "LWPOLYLINE":
            try:
                c = list(obj.Coordinates)
                d["points"] = [[round(c[i], 6), round(c[i + 1], 6)] for i in range(0, len(c), 2)]
            except Exception:
                pass
        if "rotation" in d:
            d["rotation"] = round(math.degrees(d["rotation"]), 6)
        return d

    def _iter_ms(self, layer: str | None):
        ms = self._ms
        for i in range(ms.Count):
            o = ms.Item(i)
            if layer is None or o.Layer.upper() == layer.upper():
                yield o

    async def entity_list(self, layer=None) -> CommandResult:
        def fn():
            ents = [{"type": self._dxf_type(o), "handle": o.Handle, "layer": o.Layer} for o in self._iter_ms(layer)]
            return {"entities": ents, "count": len(ents)}
        return await self._run(fn, undo=False)

    async def entity_count(self, layer=None) -> CommandResult:
        return await self._run(lambda: {"count": sum(1 for _ in self._iter_ms(layer))}, undo=False)

    async def entity_get(self, entity_id) -> CommandResult:
        return await self._run(lambda: self._describe(self._obj(entity_id)), undo=False)

    async def entity_erase(self, entity_id) -> CommandResult:
        def fn():
            o = self._obj(entity_id)
            h = o.Handle
            o.Delete()
            return f"erased {h}"
        return await self._run(fn)

    async def entity_copy(self, entity_id, dx, dy) -> CommandResult:
        def fn():
            c = self._obj(entity_id).Copy()
            c.Move(self._pt(0, 0), self._pt(dx, dy))
            return self._done(c)
        return await self._run(fn)

    async def entity_move(self, entity_id, dx, dy) -> CommandResult:
        def fn():
            o = self._obj(entity_id)
            o.Move(self._pt(0, 0), self._pt(dx, dy))
            return {"handle": o.Handle, "moved": [dx, dy]}
        return await self._run(fn)

    async def entity_rotate(self, entity_id, cx, cy, angle) -> CommandResult:
        def fn():
            o = self._obj(entity_id)
            o.Rotate(self._pt(cx, cy), math.radians(angle))
            return {"handle": o.Handle, "rotated": angle}
        return await self._run(fn)

    async def entity_scale(self, entity_id, cx, cy, factor) -> CommandResult:
        def fn():
            o = self._obj(entity_id)
            o.ScaleEntity(self._pt(cx, cy), float(factor))
            return {"handle": o.Handle, "scaled": factor}
        return await self._run(fn)

    async def entity_mirror(self, entity_id, x1, y1, x2, y2) -> CommandResult:
        # Keeps the source, like the LISP backend (MIRROR ... _N).
        return await self._run(lambda: self._done(self._obj(entity_id).Mirror(self._pt(x1, y1), self._pt(x2, y2))))

    async def entity_offset(self, entity_id, distance) -> CommandResult:
        def fn():
            res = self._obj(entity_id).Offset(float(distance))
            handles = [o.Handle for o in res]
            if handles:
                self._last_handle = handles[-1]
            return {"handle": handles[0] if handles else None, "handles": handles}
        return await self._run(fn)

    async def entity_array(self, entity_id, rows, cols, row_dist, col_dist) -> CommandResult:
        def fn():
            src = self._obj(entity_id)
            handles = []
            for r in range(int(rows)):
                for c in range(int(cols)):
                    if r == 0 and c == 0:
                        continue
                    cp = src.Copy()
                    cp.Move(self._pt(0, 0), self._pt(c * col_dist, r * row_dist))
                    handles.append(cp.Handle)
            if handles:
                self._last_handle = handles[-1]
            return {"source": src.Handle, "copies": len(handles), "handles": handles}
        return await self._run(fn)

    # ------------------------------------------------------------------
    # Layers
    # ------------------------------------------------------------------

    async def layer_list(self) -> CommandResult:
        def fn():
            out = []
            ls = self._doc.Layers
            for i in range(ls.Count):
                l = ls.Item(i)
                out.append({"name": l.Name, "color": l.color, "linetype": l.Linetype,
                            "lineweight": l.Lineweight, "on": l.LayerOn, "frozen": l.Freeze, "locked": l.Lock})
            return {"layers": out, "current": self._doc.ActiveLayer.Name}
        return await self._run(fn, undo=False)

    def _apply_layer_props(self, layer, color=None, linetype=None, lineweight=None):
        c = self._color(color)
        if c is not None:
            layer.color = c
        if linetype:
            self._ensure_linetype(linetype)
            layer.Linetype = linetype
        if lineweight not in (None, ""):
            lw = str(lineweight).strip()
            v = int(round(float(lw) * 100)) if "." in lw or float(lw) < 3 else int(lw)
            if v not in _LINEWEIGHTS:
                raise ValueError(f"Invalid lineweight '{lineweight}' (mm, e.g. 0.25, 0.50)")
            layer.Lineweight = v

    async def layer_create(self, name, color="white", linetype="CONTINUOUS") -> CommandResult:
        def fn():
            ls = self._doc.Layers
            try:
                layer = ls.Item(name)
                existed = True
            except Exception:
                layer = ls.Add(name)
                existed = False
            self._apply_layer_props(layer, color, linetype)
            return {"name": layer.Name, "color": layer.color, "linetype": layer.Linetype, "existed": existed}
        return await self._run(fn)

    async def layer_set_current(self, name) -> CommandResult:
        def fn():
            self._ensure_layer(name)
            layer = self._doc.Layers.Item(name)
            if layer.Freeze:
                layer.Freeze = False
            self._doc.ActiveLayer = layer
            return {"current_layer": layer.Name}
        return await self._run(fn)

    async def layer_set_properties(self, name, color=None, linetype=None, lineweight=None) -> CommandResult:
        def fn():
            layer = self._doc.Layers.Item(name)
            self._apply_layer_props(layer, color, linetype, lineweight)
            return {"name": layer.Name, "color": layer.color, "linetype": layer.Linetype, "lineweight": layer.Lineweight}
        return await self._run(fn)

    async def _layer_flag(self, name, prop, value) -> CommandResult:
        def fn():
            layer = self._doc.Layers.Item(name)
            if prop == "Freeze" and value and self._doc.ActiveLayer.Name.upper() == layer.Name.upper():
                raise ValueError(f"Cannot freeze the current layer '{name}'; set another layer current first")
            setattr(layer, prop, value)
            return {"name": layer.Name, prop.lower(): value}
        return await self._run(fn)

    async def layer_freeze(self, name) -> CommandResult:
        return await self._layer_flag(name, "Freeze", True)

    async def layer_thaw(self, name) -> CommandResult:
        return await self._layer_flag(name, "Freeze", False)

    async def layer_lock(self, name) -> CommandResult:
        return await self._layer_flag(name, "Lock", True)

    async def layer_unlock(self, name) -> CommandResult:
        return await self._layer_flag(name, "Lock", False)

    # ------------------------------------------------------------------
    # Blocks
    # ------------------------------------------------------------------

    def _user_blocks(self) -> list[str]:
        out = []
        bs = self._doc.Blocks
        for i in range(bs.Count):
            b = bs.Item(i)
            if not b.IsLayout and not b.IsXRef and not b.Name.startswith("*"):
                out.append(b.Name)
        return out

    async def block_list(self) -> CommandResult:
        return await self._run(lambda: {"blocks": self._user_blocks()}, undo=False)

    async def block_define(self, name, entities) -> CommandResult:
        """Define a block. Entity dicts (coordinates relative to the base point 0,0):

        LINE {x1,y1,x2,y2} · CIRCLE {cx,cy,radius} · ARC {cx,cy,radius,start_angle,end_angle (deg)}
        POLYLINE {points:[[x,y],...], closed?} · RECTANGLE {x1,y1,x2,y2} · TEXT {x,y,text,height?}
        ATTDEF {tag, x, y, height?, prompt?, default?, invisible?}
        """
        def fn():
            bs = self._doc.Blocks
            try:
                bs.Item(name)
                raise ValueError(f"Block '{name}' already exists")
            except ValueError:
                raise
            except Exception:
                pass
            blk = bs.Add(self._pt(0, 0), name)
            n = 0
            for e in entities:
                t = str(e.get("type", "LINE")).upper()
                g = lambda k, d=0.0: float(e.get(k, d))
                if t == "LINE":
                    blk.AddLine(self._pt(g("x1"), g("y1")), self._pt(g("x2"), g("y2")))
                elif t == "CIRCLE":
                    blk.AddCircle(self._pt(g("cx"), g("cy")), g("radius", 1))
                elif t == "ARC":
                    blk.AddArc(self._pt(g("cx"), g("cy")), g("radius", 1),
                               math.radians(g("start_angle")), math.radians(g("end_angle", 90)))
                elif t in ("POLYLINE", "LWPOLYLINE", "RECTANGLE"):
                    if t == "RECTANGLE":
                        pts = [[g("x1"), g("y1")], [g("x2"), g("y1")], [g("x2"), g("y2")], [g("x1"), g("y2")]]
                        closed = True
                    else:
                        pts, closed = e.get("points", []), bool(e.get("closed", False))
                    pl = blk.AddLightWeightPolyline(self._darr([c for p in pts for c in p[:2]]))
                    pl.Closed = closed
                elif t == "TEXT":
                    blk.AddText(str(e.get("text", "")), self._pt(g("x"), g("y")), g("height", 2.5))
                elif t == "ATTDEF":
                    # Mode: acAttributeModeNormal=0, acAttributeModeInvisible=1
                    mode = 1 if e.get("invisible", False) else 0
                    blk.AddAttribute(g("height", 2.5), mode, str(e.get("prompt", e.get("tag", "TAG"))),
                                     self._pt(g("x"), g("y")), str(e.get("tag", "TAG")).upper(),
                                     str(e.get("default", "")))
                else:
                    raise ValueError(f"Unsupported block entity type '{t}'")
                n += 1
            return {"block": name, "entity_count": n}
        return await self._run(fn)

    def _insert(self, name, x, y, scale, rotation, attributes):
        if name not in self._user_blocks():
            raise ValueError(f"Block '{name}' not found")
        s = float(scale or 1.0)
        ref = self._ms.InsertBlock(self._pt(x, y), name, s, s, s, math.radians(rotation or 0.0))
        set_tags = []
        if attributes and ref.HasAttributes:
            wanted = {str(k).upper(): str(v) for k, v in attributes.items()}
            for att in ref.GetAttributes():
                if att.TagString.upper() in wanted:
                    att.TextString = wanted[att.TagString.upper()]
                    set_tags.append(att.TagString)
        out = self._done(ref)
        if attributes:
            out["attributes_set"] = set_tags
        return out

    async def block_insert(self, name, x, y, scale=1.0, rotation=0.0, block_id=None) -> CommandResult:
        return await self._run(lambda: self._insert(name, x, y, scale, rotation, None))

    async def block_insert_with_attributes(self, name, x, y, scale=1.0, rotation=0.0, attributes=None) -> CommandResult:
        return await self._run(lambda: self._insert(name, x, y, scale, rotation, attributes))

    async def block_get_attributes(self, entity_id) -> CommandResult:
        def fn():
            ref = self._obj(entity_id)
            if not ref.HasAttributes:
                return {"attributes": {}}
            return {"attributes": {a.TagString: a.TextString for a in ref.GetAttributes()}}
        return await self._run(fn, undo=False)

    async def block_update_attribute(self, entity_id, tag, value) -> CommandResult:
        def fn():
            ref = self._obj(entity_id)
            for a in ref.GetAttributes() if ref.HasAttributes else []:
                if a.TagString.upper() == str(tag).upper():
                    a.TextString = str(value)
                    return {"tag": a.TagString, "value": str(value)}
            raise ValueError(f"Attribute '{tag}' not found on {entity_id}")
        return await self._run(fn)

    # ------------------------------------------------------------------
    # Annotation
    # ------------------------------------------------------------------

    async def create_text(self, x, y, text, height=2.5, rotation=0.0, layer=None) -> CommandResult:
        def fn():
            t = self._ms.AddText(str(text), self._pt(x, y), float(height))
            t.Alignment = 10  # acAlignmentMiddleCenter (same anchor as TEXT _J _M)
            t.TextAlignmentPoint = self._pt(x, y)
            t.Rotation = math.radians(rotation or 0.0)
            return self._done(t, layer)
        return await self._run(fn)

    async def create_dimension_linear(self, x1, y1, x2, y2, dim_x, dim_y) -> CommandResult:
        def fn():
            # Horizontal when the dimension line sits above/below the points, vertical otherwise.
            between_x = min(x1, x2) <= dim_x <= max(x1, x2)
            horizontal = between_x or abs(x2 - x1) >= abs(y2 - y1)
            rot = 0.0 if horizontal else math.pi / 2
            d = self._ms.AddDimRotated(self._pt(x1, y1), self._pt(x2, y2), self._pt(dim_x, dim_y), rot)
            return self._done(d, extra={"measurement": round(d.Measurement, 6)})
        return await self._run(fn)

    async def create_dimension_aligned(self, x1, y1, x2, y2, offset) -> CommandResult:
        def fn():
            L = math.hypot(x2 - x1, y2 - y1) or 1.0
            nx, ny = -(y2 - y1) / L, (x2 - x1) / L
            loc = self._pt((x1 + x2) / 2 + nx * offset, (y1 + y2) / 2 + ny * offset)
            d = self._ms.AddDimAligned(self._pt(x1, y1), self._pt(x2, y2), loc)
            return self._done(d, extra={"measurement": round(d.Measurement, 6)})
        return await self._run(fn)

    async def create_dimension_angular(self, cx, cy, x1, y1, x2, y2) -> CommandResult:
        def fn():
            a1, a2 = math.atan2(y1 - cy, x1 - cx), math.atan2(y2 - cy, x2 - cx)
            mid = (a1 + a2) / 2
            r = 0.6 * min(math.hypot(x1 - cx, y1 - cy), math.hypot(x2 - cx, y2 - cy))
            d = self._ms.AddDimAngular(self._pt(cx, cy), self._pt(x1, y1), self._pt(x2, y2),
                                       self._pt(cx + r * math.cos(mid), cy + r * math.sin(mid)))
            return self._done(d, extra={"measurement_deg": round(math.degrees(d.Measurement), 6)})
        return await self._run(fn)

    async def create_dimension_radius(self, cx, cy, radius, angle) -> CommandResult:
        def fn():
            a = math.radians(angle)
            d = self._ms.AddDimRadial(self._pt(cx, cy), self._pt(cx + radius * math.cos(a), cy + radius * math.sin(a)),
                                      float(radius) * 0.5)
            return self._done(d, extra={"measurement": round(d.Measurement, 6)})
        return await self._run(fn)

    async def create_leader(self, points, text) -> CommandResult:
        def fn():
            if len(points) < 2:
                raise ValueError("A leader needs at least 2 points")
            last = points[-1]
            h = float(self._doc.GetVariable("TEXTSIZE"))  # current default text height
            mt = self._ms.AddMText(self._pt(last[0], last[1]), 0.0, str(text))
            mt.Height = h
            flat = [c for p in points for c in (p[0], p[1], 0.0)]
            ld = self._ms.AddLeader(self._darr(flat), mt, 0)  # acLineWithArrow
            out = self._done(ld)
            out["text_handle"] = mt.Handle
            return out
        return await self._run(fn)

    # ------------------------------------------------------------------
    # View
    # ------------------------------------------------------------------

    async def zoom_extents(self) -> CommandResult:
        def fn():
            self._connect().ZoomExtents()
            return "zoomed to extents"
        return await self._run(fn, undo=False)

    async def zoom_window(self, x1, y1, x2, y2) -> CommandResult:
        def fn():
            self._connect().ZoomWindow(self._pt(x1, y1), self._pt(x2, y2))
            return "zoomed to window"
        return await self._run(fn, undo=False)
