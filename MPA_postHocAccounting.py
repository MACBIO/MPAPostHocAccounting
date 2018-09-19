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
from PyQt5.QtCore import QFileInfo
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QAction, QTreeWidgetItem, QTableWidgetItem, QFileDialog
# Initialize Qt resources from file resources.py
from .resources import *
# Import the code for the dialog
from .MPA_postHocAccounting_dialog_base import MPAPostHocAccountingDialogBase
from .MPA_postHocAccounting_dialog_targets import MPAPostHocAccountingDialogTargets
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

        # Declare instance attributes
        self.actions = []
        self.menu = u'MPA Post-Hoc Accounting'
        self.toolbar = self.iface.addToolBar(u'MPAPostHocAccounting')
        self.toolbar.setObjectName(u'MPAPostHocAccounting')

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

        # Create the dialog and keep reference
        self.dlg_base = MPAPostHocAccountingDialogBase()
        self.dlg_targets = MPAPostHocAccountingDialogTargets()

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
            text=u'MPA Post-Hoc Accounting',
            callback=self.run,
            parent=self.iface.mainWindow())

    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                u'&MPA Post-Hoc Accounting',
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar

    def run(self):
        """Run method that performs all the real work"""

        # show the window
        self.dlg_base.show()
        
        # select the MPA layer
        self.in_mpa_layer = self.dlg_base.inMPA_Layer.currentLayer()

        def set_layer_name():
            self.in_mpa_layer = self.dlg_base.inMPA_Layer.currentLayer()

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_layer_name)

        # set the mpaLayer for the field combo box
        def set_field_combo_box_layer(in_layer):
            self.dlg_base.fieldComboBox.setLayer(in_layer)

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_field_combo_box_layer)
        
        # set the MPA unique identifier field
        def set_mpa_field():
            self.inMPAfield = self.dlg_base.fieldComboBox.currentField()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_mpa_field)
        self.inMPAfield = self.dlg_base.fieldComboBox.currentField()
            
        # add polygon layers and field names to tree widget
        def set_layers():
            # add layer names and field names to analysis selection window
            layer_fields_tree = self.dlg_base.inData
            layer_fields_tree.clear()
            for layer in self.iface.mapCanvas().layers():
                if layer.name() == self.in_mpa_layer.name():
                    pass
                else:
                    tree_item = QTreeWidgetItem()
                    layer_fields_tree.addTopLevelItem(tree_item)
                    tree_item.setText(0, layer.name())
                    for field in layer.fields():
                        field_item = QTreeWidgetItem(tree_item)
                        field_item.setText(0, field.name())
            self.dlg_base.inData.expandAll()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_layers)
            
        # add selected layers and fields to processing list
        self.checkPolyDict = {}

        def tree_selection_changed():
            self.checkPolyDict = {}
            get_selected = self.dlg_base.inData.selectedItems()
            for i in get_selected:
                if i.parent():
                    field_name = i.text(0)
                    layer_name = i.parent().text(0)
                    for j in range(self.iface.mapCanvas().layerCount()):
                        layer = self.iface.mapCanvas().layer(j)
                        if layer.name() == layer_name:
                            for field in layer.fields():
                                if field.name() == field_name:
                                    self.checkPolyDict[layer_name] = {'layer': layer, 'field': field}
        self.dlg_base.inData.itemSelectionChanged.connect(tree_selection_changed)
        
        # function returns dict of dicts with area of intersection for two shapefiles
        def intersect_area(layer1, field1, layer2, field2):
            area_dict = {}
            # loop through features in first shapefile
            for feat1 in layer1.getFeatures():
                feat_dict = {}
                geom1 = feat1.geometry()
                attr1 = feat1[layer1.fields().lookupField(field1)]
                # loop through features in second shapefile
                for feat2 in layer2.getFeatures():
                    geom2 = feat2.geometry()
                    attr2 = feat2[layer2.fields().lookupField(field2)]
                    # if features intersect then write feature attr and area to shape2 dict
                    if geom2.intersects(geom1):
                        intersection = geom1.intersection(geom2)
                        int_area = intersection.area()
                        if attr2 in feat_dict.keys():
                            feat_dict[attr2] += (int_area / geom1.area())
                        else:
                            feat_dict[attr2] = (int_area / geom1.area())
                    # write shape2 dict to output dict
                area_dict[attr1] = feat_dict
            return area_dict
                
        # this part is executed after the ok button is pressed on the base window
        result_base = self.dlg_base.exec_()
        if result_base:
            table_widget = self.dlg_targets.tableWidget
            self.dlg_targets.tableWidget.clearContents()
            row = 0
            for layer in self.checkPolyDict.keys():
                table_widget.insertRow(row)
                table_item = QTableWidgetItem()
                table_item.setText(layer)
                table_widget.setItem(row, 0, table_item)
                row += 1
            for row in range(table_widget.rowCount()):
                for column in range(1, table_widget.columnCount()):
                    val = 0
                    if column == 1:
                        val = 10
                    elif column == 2:
                        val = 2
                    table_item = QTableWidgetItem()
                    table_item.setText(str(val))
                    table_widget.setItem(row, column, table_item)
                    
            self.dlg_targets.show()
            
            # display file dialog to select output spreadsheet
            self.outXLS = ''

            def out_file():
                file_dialog = QFileDialog()
                file_dialog.setOption(QFileDialog.DontConfirmOverwrite)
                file_dialog.setOption(QFileDialog.DontUseNativeDialog)
                out_name, _filter = file_dialog.getSaveFileName(file_dialog, "Output Spreadsheet", os.getenv('HOME'), "Spreadsheets (*.xls)")
                out_path = QFileInfo(out_name).absoluteFilePath()
                if not out_path.upper().endswith(".XLS"):
                    out_path = out_path + ".xls"
                if out_name:
                    self.outXLS = out_path
                    self.dlg_targets.outTable.clear()
                    self.dlg_targets.outTable.insert(out_path)
           
            # select output spreadsheet file
            self.dlg_targets.outButton.clicked.connect(out_file)
            
            # this part is executed after the ok button is pressed on the targets window
            result_target = self.dlg_targets.exec_()
            if result_target:
                # set the coverage and replication targets
                for row in range(table_widget.rowCount()):
                    layer_name = ''
                    coverage_target = 0
                    replication_target = 0
                    for column in range(table_widget.columnCount()):
                        table_item = table_widget.item(row, column)
                        if column == 0:
                            layer_name = table_item.text()
                        elif column == 1:
                            coverage_target = table_item.text()
                        elif column == 2:
                            replication_target = table_item.text()
                    self.checkPolyDict[layer_name]['coverageTarget'] = int(coverage_target)
                    self.checkPolyDict[layer_name]['replTarget'] = int(replication_target)
            
                # create a workbook
                wb = xlwt.Workbook()

                # analyse mpa size and distance to nearest MPA
                # create dict of distances to other mpas
                if self.in_mpa_layer.featureCount() > 1:
                    dist_dict = {}
                    for feat in self.in_mpa_layer.getFeatures():
                        geom = feat.geometry()
                        attr = feat.attribute(self.inMPAfield)
                        dist_list = []
                        attr_list = []
                        for test_feature in self.in_mpa_layer.getFeatures():
                            dist = geom.distance(test_feature.geometry())
                            if dist == 0.0:
                                pass
                            else:
                                dist_list.append(dist)
                                attr_list.append(test_feature.attribute(self.inMPAfield))

                        # add values to dictionary
                        mindist = min(dist_list)
                        minattr = attr_list[dist_list.index(mindist)]
                        dist_dict[attr] = [minattr, mindist]

                    # write closest mpa to workbook
                    ws = wb.add_sheet("Distances")
                    row = 0
                    # write header
                    header_cells = ["MPA ID", "Nearest MPA", "Distance (km)"]
                    for i in range(len(header_cells)):
                        ws.write(row, i, header_cells[i])
                    row += 1
                    # write values
                    for item in dist_dict.keys():
                        ws.write(row, 0, item)
                        ws.write(row, 1, dist_dict[item][0])
                        ws.write(row, 2, 111 * dist_dict[item][1])  # this is a rough conversion from DD to kilometres
                        row += 1
                
                # loop through polygon layers
                for polyName in self.checkPolyDict.keys():
                    # add a new worksheet to workbook
                    ws = wb.add_sheet(polyName)
                    row = 1
                    # define green and red styles
                    green_style = 'pattern: pattern solid, pattern_fore_colour lime, pattern_back_colour lime'
                    red_style = 'pattern: pattern solid, pattern_fore_colour rose, pattern_back_colour rose'
                    # get information from the dictionary
                    layer = self.checkPolyDict[polyName]['layer']
                    field = self.checkPolyDict[polyName]['field']
                    coverage_target = self.checkPolyDict[polyName]['coverageTarget']
                    replication_target = self.checkPolyDict[polyName]['replTarget']
                    # write header
                    header_cells = [polyName + " " + field.name(),
                                    "Coverage" + " target=" + "{0:.0f}%".format(coverage_target),
                                    "Replication" + ' target=' + str(replication_target)]
                    for i in range(len(header_cells)):
                        ws.write(0, i, header_cells[i])
                    # create list of unique IDs for polygons
                    attr_index = layer.fields().lookupField(field.name())
                    attr_list = layer.uniqueValues(attr_index)
                    # create dictionary with entry for each polygon with values of area intersecting with each MPA
                    mpa_area_per_poly = intersect_area(layer, field.name(), self.in_mpa_layer, self.inMPAfield)
                    # print report 
                    for uniqueID in attr_list:
                        sum_area = sum(mpa_area_per_poly[uniqueID].values())
                        mpa_count = str(len([PA for PA in mpa_area_per_poly[uniqueID]]))
                        print_list = [uniqueID, sum_area, mpa_count]
                        for i in range(len(print_list)):
                            attribute = print_list[i]
                            if i == 0:
                                ws.write(row, 0, attribute)
                            elif i == 1:
                                attribute = float(attribute)
                                if attribute >= coverage_target / 100.0:
                                    style_string = green_style
                                else:
                                    style_string = red_style
                                style = xlwt.easyxf(style_string, num_format_str='0%')
                                ws.write(row, 1, attribute, style)
                            elif i == 2:
                                attribute = int(attribute)
                                if attribute >= replication_target:
                                    style_string = green_style
                                else:
                                    style_string = red_style
                                style = xlwt.easyxf(style_string)
                                ws.write(row, 2, attribute, style)
                        row += 1
                    for i in range(len(header_cells)):
                        ws.col(i).width = (len(header_cells[i]) + 4) * 367
                wb.save(self.outXLS)
                if os.path.exists(self.outXLS):
                    os.system(self.outXLS)

                # clean up variables
                self.in_mpa_layer = None
                self.inMPAfield = None
                self.checkPolyDict = None