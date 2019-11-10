import sympy

test = r"\documentclass[12pt,a4paper]{article}\usepackage{array}\usepackage{booktabs}\usepackage{siunitx}\begin{document}\begin{tabular}{lS}& 354.42\\$+$& 42.1\\\hline %or \bottomrule if using the `booktabs` package& 396.52\\\end{tabular}\end{document}"

sympy.preview(test, viewer='file', filename='test2.jpg', euler=False)

sympy.preview(Eq(x + 42, y), viewer='file', filename='test2.jpg', euler=False)