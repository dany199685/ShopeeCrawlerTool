from PyQt5 import QtWidgets
import globals as Global

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Global._init ()
    Global.MainWindow.show ()
    sys.exit(app.exec_())