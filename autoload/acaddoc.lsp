;; Auto-carga del MCP de AutoCAD (puran-water) en cada dibujo.
;; Creado por la instalacion del kit de arquitectura. Borrar este archivo desactiva la auto-carga.
(if (findfile "C:/mcp/autocad-mcp/lisp-code/mcp_dispatch.lsp")
  (load "C:/mcp/autocad-mcp/lisp-code/mcp_dispatch.lsp"))
(princ)
