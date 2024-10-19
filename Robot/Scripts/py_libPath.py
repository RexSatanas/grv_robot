import os
import sys


def add_import_path():
    possible_path = [os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'engineer-tabelshik\\lib'),
                     "C:\\Projects", "C:\\Workspace\\robots\\python_robots",
                     "C:\\Robot_GVC\\Scripts\\lib", "C:\\Scripts\\lib"]
    for path in possible_path:
        if os.path.exists(path):
            libs = [folder for folder in os.listdir(path)]
            for lib in libs:
                if os.path.join(path, lib) not in sys.path and os.path.isdir(os.path.join(path, lib)):
                    sys.path.append(os.path.join(path, lib))


add_import_path()
