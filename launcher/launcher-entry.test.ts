describe('launcher entrypoint', () => {
  it('should invoke executeLauncher with the compiled launcher directory when the bin entrypoint runs', () => {
    jest.isolateModules(() => {
      const executeLauncher = jest.fn();

      jest.doMock('./run-launcher', () => ({
        executeLauncher,
      }));

      require('./launcher');

      expect(executeLauncher).toHaveBeenCalledWith(expect.any(String), process);
    });
  });
});
