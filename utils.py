from PyQt4.QtCore import *
from PyQt4.QtGui import *
from qgis.core import *
from qgis.gui import *

LAYER = 1002
FIELD = 1003

class LayerItem(QTreeWidgetItem):
    def __init__(self,parent,layer):
        QTreeWidgetItem.__init__(self,parent,[layer.name()],LAYER)
        self.layer = layer

class FieldItem(QTreeWidgetItem):
    def __init__(self,parent,layer,fieldIndex,field):
        QTreeWidgetItem.__init__(self,parent,[field.name()],FIELD)
        self.layer = layer
        self.fieldIndex = fieldIndex
        self.field = field

def fieldNames(layer):
    p = layer.dataProvider()
    fm = p.fields()
    names = []
    for f in fm:
        names.append(fm[f].name())
    # array of QStrings
    return names
    
# class Chooser(QDialog, Ui_Dialog):
    # def __init__(self, layers):

        # layerFieldsTree = self.dlg.layerFields

        # self.items = []
        # for layer in layers:
            # item = LayerItem(None, layer)
            # fields = layer.pendingFields()
            # for k,v in enumerate(fields):
                # jtem = FieldItem(item, layer, k, v)
            # self.items.append(item)
            # print item
        # layerFieldsTree.insertTopLevelItems(0, self.items);
        # for item in self.items:
            # layerFieldsTree.expandItem(item)
            # print item

    # def clickedItem(self, item, column):
        # if item.type() == FIELD:
            # self.buttonBox.button(QDialogButtonBox.Ok).setEnabled(True)
        # else:
            # self.buttonBox.button(QDialogButtonBox.Ok).setEnabled(False)
            
        # self.clicked=item

    # def clickedItemDouble(self,item,column):
        # if item.type() == FIELD:
            # self.clicked = item
            # self.accept()
        # else:
            # self.buttonBox.button(QDialogButtonBox.Ok).setEnabled(False)