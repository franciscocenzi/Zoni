from qgis.PyQt.QtCore import QObject, QEvent, Qt

class EnterKeyFilter(QObject):

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
            self.callback()
            return True
        return False