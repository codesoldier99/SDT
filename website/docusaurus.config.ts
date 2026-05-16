import {themes as prismThemes} from 'prism-react-renderer';
import type {Config} from '@docusaurus/types';
import type * as Preset from '@docusaurus/preset-classic';

// This runs in Node.js - Don't use client-side code here (browser APIs, JSX...)

const repository = process.env.GITHUB_REPOSITORY ?? 'your-org/SDT';
const [organizationName, projectName] = repository.split('/');
const isGitHubActions = process.env.GITHUB_ACTIONS === 'true';

const config: Config = {
  title: '传感器与检测技术',
  tagline: '具身感知工程版课程网站与教学资源仓库',
  favicon: 'img/favicon.ico',

  // Future flags, see https://docusaurus.io/docs/api/docusaurus-config#future
  future: {
    v4: true, // Improve compatibility with the upcoming Docusaurus v4
  },

  // Set the production url of your site here
  url: isGitHubActions ? `https://${organizationName}.github.io` : 'http://localhost',
  // Set the /<baseUrl>/ pathname under which your site is served
  // For GitHub pages deployment, it is often '/<projectName>/'
  baseUrl: isGitHubActions ? `/${projectName}/` : '/',

  // GitHub pages deployment config.
  // If you aren't using GitHub pages, you don't need these.
  organizationName,
  projectName,

  onBrokenLinks: 'throw',

  // Even if you don't use internationalization, you can use this field to set
  // useful metadata like html lang. For example, if your site is Chinese, you
  // may want to replace "en" with "zh-Hans".
  i18n: {
    defaultLocale: 'zh-Hans',
    locales: ['zh-Hans'],
  },

  presets: [
    [
      'classic',
      {
        docs: {
          sidebarPath: './sidebars.ts',
        },
        blog: {
          routeBasePath: 'announcements',
          blogTitle: '课程公告',
          blogDescription: '课程更新、开课提醒和版本发布公告',
          blogSidebarTitle: '全部公告',
          blogSidebarCount: 'ALL',
          showReadingTime: true,
          feedOptions: {
            type: ['rss', 'atom'],
            xslt: true,
          },
          // Useful options to enforce blogging best practices
          onInlineTags: 'warn',
          onInlineAuthors: 'warn',
          onUntruncatedBlogPosts: 'warn',
        },
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Preset.Options,
    ],
  ],

  themeConfig: {
    // Replace with your project's social card
    image: 'img/docusaurus-social-card.jpg',
    colorMode: {
      respectPrefersColorScheme: true,
    },
    navbar: {
      title: '传感器与检测技术',
      logo: {
        alt: '传感器与检测技术',
        src: 'img/logo.svg',
      },
      items: [
        {
          to: '/docs/intro',
          position: 'left',
          label: '课程总览',
        },
        {to: '/docs/syllabus', label: 'Syllabus', position: 'left'},
        {to: '/docs/schedule', label: 'Schedule', position: 'left'},
        {to: '/docs/lectures', label: 'Lectures', position: 'left'},
        {to: '/docs/labs', label: 'Labs', position: 'left'},
        {to: '/docs/resources', label: 'Resources', position: 'left'},
        {to: '/docs/faq', label: 'FAQ', position: 'left'},
        {to: '/announcements', label: 'Announcements', position: 'left'},
        {
          href: `https://github.com/${organizationName}/${projectName}`,
          label: 'GitHub',
          position: 'right',
        },
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {
          title: '课程文档',
          items: [
            {
              label: '课程总览',
              to: '/docs/intro',
            },
            {
              label: '授课计划',
              to: '/docs/syllabus',
            },
            {
              label: '教学进度',
              to: '/docs/schedule',
            },
          ],
        },
        {
          title: '教学支持',
          items: [
            {
              label: '实验安排',
              to: '/docs/labs',
            },
            {
              label: '参考资料',
              to: '/docs/resources',
            },
            {
              label: 'FAQ',
              to: '/docs/faq',
            },
          ],
        },
        {
          title: '更多',
          items: [
            {
              label: '课程公告',
              to: '/announcements',
            },
            {
              label: '版本归档',
              to: '/docs/archive',
            },
            {
              label: 'GitHub',
              href: `https://github.com/${organizationName}/${projectName}`,
            },
          ],
        },
      ],
      copyright: `Copyright © ${new Date().getFullYear()} 传感器与检测技术课程团队。Built with Docusaurus.`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.vsDark,
    },
  } satisfies Preset.ThemeConfig,
};

export default config;
