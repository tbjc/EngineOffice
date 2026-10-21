# Sitio de documentación de OfficeEngine

Sitio de [Docusaurus](https://docusaurus.io) con la documentación de uso de la biblioteca: **índice,
Word, Excel, PDF y formatos heredados**, con un ejemplo por función.

**Publicado en**: <https://tbjc.github.io/EngineOffice/>

## Ponerlo en marcha

```bash
npm install
npm start        # servidor local en http://localhost:3000/EngineOffice/
npm run build    # sitio estático en build/
```

El sitio se publica solo: el flujo `deploy-docs` de `.github/workflows/ci.yml` lo sube a GitHub
Pages en cada `push` a `main`. Las rutas del sitio usan `baseUrl: '/EngineOffice/'`, que es el
nombre del repositorio: si el repositorio cambia de nombre, hay que cambiar también el `baseUrl`.

## Estructura

```
docs/            Contenido (Markdown). Una página por función, en el orden del índice lateral.
sidebars.ts      El orden de la navegación.
src/css/         Estilos mínimos sobre el tema clásico.
verify-examples.ps1  Compila los ejemplos del sitio contra la biblioteca real.
```

## Verificar los ejemplos

Cada bloque ```csharp de las páginas se compila contra los proyectos de la biblioteca, en un
proyecto temporal. Un ejemplo que no compila rompe la verificación.

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File docs-site/verify-examples.ps1
```

## Reglas del contenido

- Un ejemplo por página, **completo** (con sus `using`) y **ejecutable**.
- Se documenta lo que la biblioteca hace, no lo que podría llegar a hacer. Cuando algo no se
  modela, se dice en la sección de notas.
- Las limitaciones van explícitas: sin comparaciones con Microsoft Office, sin paridad con otras
  librerías.
