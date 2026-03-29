import * as packageJson from '../package.json';

describe('package.json publish settings', () => {
  it('should expose the npx entrypoint from dist/launcher.js', () => {
    expect(packageJson.bin).toEqual({
      'excel-cli': 'dist/launcher.js',
    });
  });

  it('should publish the scoped package as public', () => {
    expect(packageJson.publishConfig).toEqual({
      access: 'public',
    });
  });

  it('should include the dist directory in the published package', () => {
    expect(packageJson.files).toContain('dist/');
  });

  it('should run jest in-band from the project test script', () => {
    expect(packageJson.scripts.test).toBe('jest --runInBand');
  });
});
