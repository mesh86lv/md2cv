{
  pkgs ? import <nixpkgs> { },
}:

let
  python = pkgs.python3.withPackages (ps: [
    ps.python-docx
    ps.pytest
  ]);
  projectDir = toString ./.;
in
pkgs.mkShell {
  packages = [
    python
    (pkgs.writeShellScriptBin "md2cv" ''
      exec ${python}/bin/python ${projectDir}/md2cv.py "$@"
    '')
  ];
}
