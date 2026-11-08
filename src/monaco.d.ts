// monaco-editor ships types only for its root entry. The editor-core entry
// used by the manifest editor exports the same API.
declare module 'monaco-editor/esm/vs/editor/edcore.main' {
  export * from 'monaco-editor';
}
