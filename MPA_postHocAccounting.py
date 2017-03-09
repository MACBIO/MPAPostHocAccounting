# -*- coding: utf-8 -*-
"""
/***************************************************************************
 MPAPostHocAccounting
                                 A QGIS plugin
 This plugin checks how your MPAs meet your placement objectives
                              -------------------
        begin                : 2017-03-02
        git sha              : $Format:%H$
        copyright            : (C) 2017 by Jonah Sullivan
        email                : jonahsullivan79@gmail.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from qgis.core import *
from qgis.gui import *
from qgis.utils import *
# Initialize Qt resources from file resources.py
import resources
# Import the code for the dialog
from MPA_postHocAccounting_dialog import MPAPostHocAccountingDialog
import os
import xlwt

class MPAPostHocAccounting:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'MPAPostHocAccounting_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)


        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&MPA Post-Hoc Accounting')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'MPAPostHocAccounting')
        self.toolbar.setObjectName(u'MPAPostHocAccounting')

    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('MPAPostHocAccounting', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):

        # Create the dialog (after translation) and keep reference
        self.dlg = MPAPostHocAccountingDialog()

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""
        icon_path = ':/plugins/MPAPostHocAccounting/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u'MPA Post-Hoc Accounting'),
            callback=self.run,
            parent=self.iface.mainWindow())

    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&MPA Post-Hoc Accounting'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar

    def run(self):
        """Run method that performs all the real work"""
        # find all of the layers in the map
        layers = []
        for i in range(iface.mapCanvas().layerCount()):
            layer = iface.mapCanvas().layer(i)
            if layer.type() == layer.VectorLayer:
                if layer.geometryType() == QGis.Polygon:
                    layers.append(layer)
       
        # list of layer names
        lyrNameList = [layer.name() for layer in layers]
        
        # clear the MPA layer dropdown
        self.dlg.inMPA_Layer.clear()
        
        # add layer names to MPA layer dropdown
        self.dlg.inMPA_Layer.addItem('select MPA layer')
        self.dlg.inMPA_Layer.addItems(lyrNameList)
        
        # set the coverage and replication target defaults
        coverageSpinBox = self.dlg.coverageSpinBox
        replSpinBox = self.dlg.replSpinBox
        coverageSpinBox.setValue(10)
        replSpinBox.setValue(2)
        
        # show the window
        self.dlg.show()
        
        # select the MPA layer
        self.inMPAlayer = QgsVectorLayer()
        def setMPAlayer():
            inMPAlayerName = self.dlg.inMPA_Layer.currentText()
            for i in range(iface.mapCanvas().layerCount()):
                    layer = iface.mapCanvas().layer(i)
                    if layer.name() == inMPAlayerName:
                        self.inMPAlayer = layer
                        # select MPA unique field next
                        fieldNameList = [field.name() for field in self.inMPAlayer.fields()]
                        self.dlg.inMPA_Field.clear()
                        self.dlg.inMPA_Field.addItems(fieldNameList)
                        setMPAfield(layer)
        self.dlg.inMPA_Layer.currentIndexChanged.connect(setMPAlayer)
        
        # set the MPA field
        self.inMPAfield = QgsField()
        def setMPAfield(layer):
            inMPAfieldName = self.dlg.inMPA_Field.currentText()
            for field in layer.pendingFields():
                if field.name() == inMPAfieldName:
                    self.inMPAfield = field
            setLayers(layers)
            
        # add polygon layers and field names to tree widget
        def setLayers(layers):
            # add layer names and field names to analysis selection window
            layerFieldsTree = self.dlg.inData
            layerFieldsTree.clear()
            for layer in layers:
                if layer.name() == self.inMPAlayer.name():
                    pass
                else:
                    #item = utils.LayerItem(None, layer)
                    treeItem = QTreeWidgetItem()
                    layerFieldsTree.addTopLevelItem(treeItem)
                    treeItem.setText(0, layer.name())
                    fields = layer.pendingFields()
                    for field in fields:
                        fieldItem = QTreeWidgetItem(treeItem)
                        fieldItem.setText(0,field.name())
                    layerFieldsTree.setItemExpanded(treeItem, True)
            
        # add selected layers and fields to processing list
        self.checkPolyDict = {}
        def treeSelectionChanged():
            self.checkPolyDict = {}
            getSelected = self.dlg.inData.selectedItems()
            for i in getSelected:
                if i.parent():
                    fieldName = i.text(0)
                    layerName = i.parent().text(0)
                    for i in range(iface.mapCanvas().layerCount()):
                        layer = iface.mapCanvas().layer(i)
                        if layer.name() == layerName:
                            for field in layer.pendingFields():
                                if field.name() == fieldName:
                                    self.checkPolyDict[layerName] = {'layer': layer, 'field': field}
        self.dlg.inData.itemSelectionChanged.connect(treeSelectionChanged)  
        
        # display file dialog to select output spreadsheet
        self.outXLS = ''
        def outFile():
            fileDialog = QFileDialog()
            fileDialog.setConfirmOverwrite(False)
            outName = fileDialog.getSaveFileName(fileDialog, "Output Spreadsheet",".", "Spreadsheets (*.xls)")
            outPath = QFileInfo(outName).absoluteFilePath()
            if not outPath.upper().endswith(".XLS"):
                outPath = outPath + ".xls"
            if outName:
                self.outXLS = outPath
                self.dlg.outTable.clear()
                self.dlg.outTable.insert(outPath)
           
        # select output spreadsheet file
        self.dlg.outButton.clicked.connect(outFile)  
        
        # function returns dict of dicts with area of intersection for two shapefiles
        def intersectArea(layer1, field1, layer2, field2):
            areaDict = {}
            # loop through features in first shapefile
            for feat1 in layer1.getFeatures():
                featDict = {}
                geom1 = feat1.geometry()
                attr1 = feat1[layer1.fieldNameIndex(field1.name())]
                # loop through features in second shapefile
                for feat2 in layer2.getFeatures():
                    geom2 = feat2.geometry()
                    attr2 = feat2[layer2.fieldNameIndex(field2.name())]
                    # if features intersect then write feature attr and area to shape2 dict
                    if geom2.intersects(geom1):
                        intersection = geom1.intersection(geom2)
                        intArea = intersection.area()
                        if attr2 in featDict.keys():
                            featDict[attr2] += (intArea / geom1.area())
                        else:
                            featDict[attr2] = (intArea / geom1.area())
                    # write shape2 dict to output dict
                areaDict[attr1] = featDict
            return areaDict
                
        # this part is executed after the ok button is pressed
        result = self.dlg.exec_()
        if result:
            # set the coverage and replication target variables
            self.coverageTarget = coverageSpinBox.value()
            self.replTarget = replSpinBox.value()
            # create a workbook
            wb = xlwt.Workbook()
            # loop through polygon layers
            for polyName in self.checkPolyDict.keys():
                # add a new worksheet to workbook
                ws = wb.add_sheet(polyName)
                row = 1
                # define green and red styles
                greenStyle = 'pattern: pattern solid, pattern_fore_colour lime, pattern_back_colour lime'
                redStyle = 'pattern: pattern solid, pattern_fore_colour rose, pattern_back_colour rose'
                # get information from the dictionary
                layer = self.checkPolyDict[polyName]['layer']
                field = self.checkPolyDict[polyName]['field']
                # write header
                headerCells = [polyName + " " + field.name(),
                               "Coverage" + " target=" + "{0:.0f}%".format(self.coverageTarget),
                               "Replication" + ' target=' + str(self.replTarget)]
                for i in range(len(headerCells)):
                    ws.write(0, i, headerCells[i])
                # create list of unique IDs for polygons
                attrIndex = layer.fieldNameIndex(field.name())
                attrList = layer.uniqueValues(attrIndex)
                # create dictionary with entry for each polygon with values of area intersecting with each MPA
                mpaAreaPerPoly = intersectArea(layer, field, self.inMPAlayer, self.inMPAfield)
                # print report 
                for uniqueID in attrList:
                    sumArea = sum(mpaAreaPerPoly[uniqueID].values())
                    mpaCount = str(len([PA for PA in mpaAreaPerPoly[uniqueID]]))
                    printList = [uniqueID, sumArea, mpaCount]
                    for attribute in printList:
                        if attribute == uniqueID:
                            try:
                                attribute = int(uniqueID)
                            except:
                                pass
                            ws.write(row, 0, attribute)
                        elif attribute == sumArea:
                            attribute = float(attribute)
                            if attribute >= self.coverageTarget/100.0:
                                styleString = greenStyle
                            else:
                                styleString = redStyle
                            style = xlwt.easyxf(styleString, num_format_str='0%')
                            ws.write(row, 1, attribute, style)
                        elif attribute == mpaCount:
                            attribute = int(attribute)
                            if attribute >= self.replTarget:
                                styleString = greenStyle
                            else:
                                styleString = redStyle
                            style = xlwt.easyxf(styleString)
                            ws.write(row, 2, attribute, style)
                    row += 1
                for i in range(len(headerCells)):
                    ws.col(i).width = (len(headerCells[i]) + 4) * 367
            wb.save(self.outXLS)
            if os.path.exists(self.outXLS):
                os.system(self.outXLS)