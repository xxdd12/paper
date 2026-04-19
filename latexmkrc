$pdf_mode = 5;

# Keep the engine and diagnostics aligned with the VS Code recipe so
# command-line latexmk and editor-triggered builds behave the same way.
$xelatex = 'xelatex -synctex=1 -interaction=nonstopmode -file-line-error %O %S';
