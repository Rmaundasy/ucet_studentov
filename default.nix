with import <nixpkgs> {};

( pkgs.python39.buildEnv.override  {
extraLibs = with pkgs.python39Packages; [ matplotlib openpyxl tkinter ];
}).env
