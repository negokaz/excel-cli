import * as path from 'path';

type BinaryTarget = {
  platform: NodeJS.Platform;
  arch: string;
  distributionDir: string;
  binaryFileName: string;
};

const SUPPORTED_TARGETS: BinaryTarget[] = [
  { platform: 'win32',  arch: 'x64',   distributionDir: 'excel-cli_windows_amd64_v1',   binaryFileName: 'excel-cli.exe' },
  { platform: 'win32',  arch: 'arm64', distributionDir: 'excel-cli_windows_arm64_v8.0', binaryFileName: 'excel-cli.exe' },
  { platform: 'darwin', arch: 'x64',   distributionDir: 'excel-cli_darwin_amd64_v1',    binaryFileName: 'excel-cli' },
  { platform: 'darwin', arch: 'arm64', distributionDir: 'excel-cli_darwin_arm64_v8.0',  binaryFileName: 'excel-cli' },
  { platform: 'linux',  arch: 'x64',   distributionDir: 'excel-cli_linux_amd64_v1',     binaryFileName: 'excel-cli' },
  { platform: 'linux',  arch: 'arm64', distributionDir: 'excel-cli_linux_arm64_v8.0',   binaryFileName: 'excel-cli' },
];

function getTarget(platform: NodeJS.Platform, arch: string): BinaryTarget {
  const target = SUPPORTED_TARGETS.find((entry) => entry.platform === platform && entry.arch === arch);

  if (target === undefined) {
    throw new Error(`Unsupported platform: ${platform}_${arch} (platform=${platform}, arch=${arch})`);
  }

  return target;
}

export function resolveBinaryPath(
  platform: NodeJS.Platform,
  arch: string,
  baseDir: string,
  existsSync: (filePath: string) => boolean,
): string {
  const target = getTarget(platform, arch);
  const binaryPath = path.resolve(baseDir, target.distributionDir, target.binaryFileName);

  if (!existsSync(binaryPath)) {
    throw new Error(`Binary not found: ${binaryPath}`);
  }

  return binaryPath;
}
