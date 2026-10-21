import type { SidebarsConfig } from '@docusaurus/plugin-content-docs';

const sidebars: SidebarsConfig = {
  lateral: [
    'index',
    {
      type: 'category',
      label: 'Word',
      collapsed: false,
      items: [
        'word/crear-y-guardar',
        'word/parrafos-y-runs',
        'word/formato-de-caracter',
        'word/formato-de-parrafo',
        'word/tabulaciones-y-saltos',
        'word/estilos',
        'word/listas',
        'word/tablas',
        'word/combinacion-de-celdas',
        'word/secciones',
        'word/encabezados-y-pies',
        'word/numeracion-de-paginas',
        'word/imagenes',
        'word/hipervinculos',
        'word/busqueda-y-reemplazo',
        'word/propiedades-del-documento',
      ],
    },
    {
      type: 'category',
      label: 'Excel',
      collapsed: false,
      items: [
        'excel/crear-y-guardar',
        'excel/celdas',
        'excel/rangos',
        'excel/fechas-y-formatos-numericos',
        'excel/estilos',
        'excel/formulas',
        'excel/nombres',
        'excel/validacion-de-datos',
        'excel/autofiltros',
        'excel/tablas-estructuradas',
        'excel/graficos',
        'excel/imagenes',
        'excel/hipervinculos-y-notas',
        'excel/vista-de-la-hoja',
        'excel/impresion',
        'excel/cadenas-compartidas',
        'excel/matriz-de-formulas',
      ],
    },
    {
      type: 'category',
      label: 'PDF',
      collapsed: false,
      items: [
        'pdf/convertir-a-pdf',
        'pdf/adaptador-libreoffice',
        'pdf/errores-y-limites',
      ],
    },
    {
      type: 'category',
      label: 'Formatos heredados',
      collapsed: false,
      items: [
        'legacy/convertir-formatos-heredados',
      ],
    },
  ],
};

export default sidebars;
