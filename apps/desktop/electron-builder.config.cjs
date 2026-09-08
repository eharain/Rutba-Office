// What the installers contain, and what the operating system is told about it.
//
// A JavaScript config rather than YAML for one reason: the file associations
// come from the same registry the Open dialog and the launcher read, so the
// suite cannot claim a format on the download page, offer it in the dialog, and
// then fail to register it with Windows.

const path = require('node:path');

module.exports = async () => {
  const { fileAssociations } = await import('@rutba/office-formats/registry');
  const year = new Date().getFullYear();

  return {
    appId: 'co.techstyle.rutba.office',
    productName: 'Rutba Office',
    copyright: `Copyright © ${year} Tech Style Ltd — AGPL-3.0-or-later`,
    artifactName: 'Rutba-Office-${version}-${os}-${arch}.${ext}',

    directories: {
      output: path.resolve(__dirname, 'release'),
      buildResources: path.resolve(__dirname, 'resources'),
    },

    // Only what runs. The renderer is one bundle, the main process is four
    // files, and the engines come in through node_modules.
    files: [
      'main/**/*',
      'build/out/**/*',
      'resources/icon.png',
      'resources/icon.ico',
      'resources/apps/*.ico',
      'package.json',
      '!**/*.map',
      '!**/test/**',
      '!**/tests/**',
      '!**/*.md',
    ],

    asar: true,
    // The taskbar reads an app's icon straight from disk, so the tiles sit
    // beside the archive rather than inside it.
    asarUnpack: ['resources/apps/*.ico'],
    // The engines read their own sources at runtime through the package
    // exports; nothing needs unpacking, and an asar makes cold start faster.
    compression: 'maximum',
    removePackageScripts: true,
    npmRebuild: false,

    // Each type wears the icon of the app that opens it — a Word document
    // looks like a document, a workbook like a grid — rather than the
    // suite's one mark on everything. The icons come from
    // build/make-file-icons.js, one per family, into resources/filetypes.
    // The owning app decides the icon and is then dropped: electron-builder
    // validates each association against its own schema and refuses a config
    // carrying a property it does not know, which is a build that never runs.
    fileAssociations: fileAssociations().map(({ app: owner, ...a }) => ({
      ...a,
      icon: path.resolve(__dirname, 'resources', 'filetypes', `${a.ext === 'pdf' ? 'pdf' : owner}.ico`),
    })),


    protocols: [{ name: 'Rutba Office', schemes: ['rutba-office'] }],

    win: {
      target: [
        { target: 'nsis', arch: ['x64'] },
        { target: 'portable', arch: ['x64'] },
      ],
      icon: path.resolve(__dirname, 'resources/icon.ico'),
      // Unsigned: this is a free download, and a certificate is a later
      // decision. SmartScreen will warn until one is bought, which is honest.
      signAndEditExecutable: true,
    },

    nsis: {
      oneClick: false,
      perMachine: false,
      allowToChangeInstallationDirectory: true,
      allowElevation: true,
      createDesktopShortcut: true,
      createStartMenuShortcut: true,
      shortcutName: 'Rutba Office',
      uninstallDisplayName: 'Rutba Office',
      deleteAppDataOnUninstall: false,
      menuCategory: false,
      installerIcon: path.resolve(__dirname, 'resources/icon.ico'),
      uninstallerIcon: path.resolve(__dirname, 'resources/icon.ico'),
      // One copy on a machine. electron-builder replaces the previous copy in
      // the scope being installed to; this script also removes one in the
      // other scope (per-user beside per-machine), whatever directory it is
      // in, and forgets a registry entry whose uninstaller is gone.
      include: path.resolve(__dirname, 'resources/installer.nsh'),
    },


    portable: {
      artifactName: 'Rutba-Office-${version}-portable.exe',
      requestExecutionLevel: 'user',
    },

    mac: {
      target: [
        { target: 'dmg', arch: ['x64', 'arm64'] },
        { target: 'zip', arch: ['x64', 'arm64'] },
      ],
      category: 'public.app-category.productivity',
      icon: path.resolve(__dirname, 'resources/icon.png'),
      darkModeSupport: true,
      hardenedRuntime: false,
      identity: null,
      extendInfo: {
        NSDocumentsFolderUsageDescription: 'Rutba Office opens and saves your documents.',
        NSDesktopFolderUsageDescription: 'Rutba Office opens and saves your documents.',
        NSDownloadsFolderUsageDescription: 'Rutba Office opens files you have downloaded.',
        NSPhotoLibraryUsageDescription: 'Rutba Pictures shows your photographs.',
      },
    },

    dmg: {
      title: 'Rutba Office ${version}',
      contents: [
        { x: 140, y: 200, type: 'file' },
        { x: 400, y: 200, type: 'link', path: '/Applications' },
      ],
    },

    linux: {
      target: [
        { target: 'AppImage', arch: ['x64'] },
        { target: 'deb', arch: ['x64'] },
      ],
      category: 'Office',
      icon: path.resolve(__dirname, 'resources/icon.png'),
      synopsis: 'Free office suite: mail, documents, worksheets, presentations and media',
      description:
        'Rutba Office is a free, open-source desktop suite: Mail, Word, Worksheets, Presentation, ' +
        'Pictures, Image and Video. It opens Microsoft and OpenDocument files, and reads the mail ' +
        'archives other clients leave behind — including Outlook .pst and .ost. It works with no ' +
        'network connection.',
      // electron-builder 26 nests the .desktop keys under `entry`.
      desktop: {
        entry: {
          StartupWMClass: 'rutba-office',
          Keywords: 'office;word;spreadsheet;presentation;mail;email;pst;image;video;',
        },
      },
    },

    deb: { depends: ['libgtk-3-0', 'libnotify4', 'libnss3', 'libxss1', 'libxtst6', 'xdg-utils', 'libatspi2.0-0'] },

    // The update feed. electron-builder writes latest.yml beside the
    // installers, and electron-updater reads it from the release.
    publish: [{ provider: 'github', owner: 'eharain', repo: 'Rutba-Office', releaseType: 'release' }],
  };
};
