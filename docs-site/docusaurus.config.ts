import { themes as prismThemes } from 'prism-react-renderer';
import type { Config } from '@docusaurus/types';
import type * as Preset from '@docusaurus/preset-classic';

const config: Config = {
  title: 'OfficeEngine',
  tagline: 'Crear, leer, modificar, guardar y convertir documentos de Word y Excel',
  favicon: 'img/favicon.svg',

  url: 'https://tbjc.github.io',
  baseUrl: '/EngineOffice/',

  organizationName: 'tbjc',
  projectName: 'EngineOffice',

  onBrokenLinks: 'throw',

  markdown: {
    hooks: {
      onBrokenMarkdownLinks: 'throw',
    },
  },

  i18n: {
    defaultLocale: 'es',
    locales: ['es'],
  },

  presets: [
    [
      'classic',
      {
        docs: {
          sidebarPath: './sidebars.ts',
          routeBasePath: '/',
        },
        blog: false,
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Preset.Options,
    ],
  ],

  themeConfig: {
    navbar: {
      title: 'OfficeEngine',
      items: [
        {
          type: 'docSidebar',
          sidebarId: 'lateral',
          position: 'left',
          label: 'Documentación',
        },
      ],
    },
    footer: {
      style: 'dark',
      copyright: `OfficeEngine — documentación de la biblioteca.`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
      additionalLanguages: ['bash', 'json', 'csharp'],
    },
  } satisfies Preset.ThemeConfig,
};

export default config;
