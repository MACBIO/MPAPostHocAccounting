# -*- coding: utf-8 -*-
"""
/***************************************************************************
 MPAPostHocAccounting
                                 A QGIS plugin
 This plugin checks how your MPAs meet your placement objectives
                              -------------------
        begin                : 2018-09-24
        git sha              : $Format:%H$
        copyright            : (C) 2018 by Jonah Sullivan
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
from PyQt5.QtWidgets import QAction, \
    QTreeWidgetItem, \
    QTableWidgetItem, \
    QFileDialog, \
    QTreeWidgetItemIterator, \
    QHeaderView
# Initialize Qt resources from file resources.py
from .resources import *
# Import the code for the dialog
from .MPA_postHocAccounting_dialog_base import MPAPostHocAccountingDialogBase
from .MPA_postHocAccounting_dialog_targets import MPAPostHocAccountingDialogTargets
import os
from qgis.core import QgsDistanceArea, QgsCoordinateReferenceSystem, QgsCoordinateTransformContext

try:
    # use system version of xlwt
    import xlwt
except ImportError:
    # use version of xlwt distributed with plugin
    import site
    import os

    # this will get the path for this file and add it to the system PATH
    # so the xlwt folder can be found
    site.addsitedir(os.path.abspath(os.path.dirname(__file__)))
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

        # initialise and clear variables
        self.out_xls = None
        self.in_mpa_field = None
        self.in_map_layer = None
        self.check_poly_dict = None
        self.dlg_base = None
        self.dlg_targets = None

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
        
        # clear old variables
        self.dlg_base.fieldComboBox.setLayer(None)
        # self.dlg_base.inMPA_Layer.clear()
        # self.dlg_base.fieldComboBox.clear()
        iterator = QTreeWidgetItemIterator(self.dlg_base.inData, QTreeWidgetItemIterator.All)
        while iterator.value():
            iterator.value().takeChildren()
            iterator += 1
        i = self.dlg_base.inData.topLevelItemCount()
        while i > -1:
            self.dlg_base.inData.takeTopLevelItem(i)
            i -= 1

        # show the window
        self.dlg_base.show()
        
        # select the MPA layer
        self.in_map_layer = self.dlg_base.inMPA_Layer.currentLayer()

        def set_layer_name():
            self.in_map_layer = self.dlg_base.inMPA_Layer.currentLayer()

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_layer_name)

        # set the mpaLayer for the field combo box
        def set_field_combo_box_layer(in_layer):
            self.dlg_base.fieldComboBox.setLayer(in_layer)

        self.dlg_base.inMPA_Layer.layerChanged.connect(set_field_combo_box_layer)
        
        # set the MPA unique identifier field
        def set_mpa_field():
            self.in_mpa_field = self.dlg_base.fieldComboBox.currentField()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_mpa_field)
        self.in_mpa_field = self.dlg_base.fieldComboBox.currentField()
            
        # add polygon layers and field names to tree widget
        def set_layers():
            # add layer names and field names to analysis selection window
            layer_fields_tree = self.dlg_base.inData
            layer_fields_tree.clear()
            for map_layer in self.iface.mapCanvas().layers():
                if map_layer.name() == self.in_map_layer.name():
                    pass
                else:
                    tree_item = QTreeWidgetItem()
                    layer_fields_tree.addTopLevelItem(tree_item)
                    tree_item.setText(0, map_layer.name())
                    for layer_field in map_layer.fields():
                        field_item = QTreeWidgetItem(tree_item)
                        field_item.setText(0, layer_field.name())
            self.dlg_base.inData.expandAll()

        self.dlg_base.fieldComboBox.fieldChanged.connect(set_layers)
            
        # add selected layers and fields to processing list
        self.check_poly_dict = {}

        def tree_selection_changed():
            self.check_poly_dict = {}
            get_selected = self.dlg_base.inData.selectedItems()
            for selected_item in get_selected:
                if selected_item.parent():
                    field_name = selected_item.text(0)
                    selected_layer_name = selected_item.parent().text(0)
                    for j in range(self.iface.mapCanvas().layerCount()):
                        selected_layer = self.iface.mapCanvas().layer(j)
                        if selected_layer.name() == selected_layer_name:
                            for layer_field in selected_layer.fields():
                                if layer_field.name() == field_name:
                                    self.check_poly_dict[selected_layer_name] = {'layer': selected_layer, 'field': layer_field}
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
            table_widget.setRowCount(0)
            row = 0
            for layer in self.check_poly_dict.keys():
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
                    
            # resize columns to fit contents
            header = table_widget.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.Stretch)
            header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

            # clear the output dialog box before showing the
            self.dlg_targets.outTable.clear()

            # show the targets dialog box
            self.dlg_targets.show()
            
            # display file dialog to select output spreadsheet
            def out_file():
                out_name, _ = QFileDialog.getSaveFileName(None,
                                                          "Output Spreadsheet",
                                                          os.getenv('HOME'),
                                                          "Spreadsheets (*.xls)")
                out_path = QFileInfo(out_name).absoluteFilePath()
                if not out_path.upper().endswith(".XLS"):
                    out_path = out_path + ".xls"
                if out_name:
                    self.out_xls = out_path
                    self.dlg_targets.outTable.setText(out_path)
           
            # select output spreadsheet file when browse button is clicked
            self.dlg_targets.outButton.clicked.connect(out_file)

            # this part is executed after the ok button is pressed on the targets window
            result_target = self.dlg_targets.exec_()
            if result_target:
                # set the coverage and replication targets
                for row in range(table_widget.rowCount()):
                    layer_name = ''
                    coverage_target = 0
                    repl_target = 0
                    for column in range(table_widget.columnCount()):
                        table_item = table_widget.item(row, column)
                        if column == 0:
                            layer_name = table_item.text()
                        elif column == 1:
                            coverage_target = table_item.text()
                        elif column == 2:
                            repl_target = table_item.text()
                    self.check_poly_dict[layer_name]['coverageTarget'] = int(coverage_target)
                    self.check_poly_dict[layer_name]['replTarget'] = int(repl_target)
            
                # create a workbook
                wb = xlwt.Workbook()

                # analyse mpa size and distance to nearest MPA
                # create dict of distances to other mpas
                if self.in_map_layer.featureCount() > 1:
                    dist_dict = {}
                    for feat in self.in_map_layer.getFeatures():
                        geom = feat.geometry()
                        attr = feat.attribute(self.in_mpa_field)
                        dist_list = []
                        attr_list = []
                        for test_feature in self.in_map_layer.getFeatures():
                            dist = geom.distance(test_feature.geometry())
                            if dist == 0.0:
                                pass
                            else:
                                # vertex in test feature closest to centroid of feature
                                closest_vertex = geom.closestVertex(test_feature.geometry().centroid().asPoint())
                                # vertex in feature closest to centroid of test feature
                                closest_vertex_test = test_feature.geometry().closestVertex(geom.centroid().asPoint())
                                # tool to measure distance between points (in metres)
                                d = QgsDistanceArea()
                                d.setEllipsoid('WGS84')
                                canvas_auth_id = self.iface.mapCanvas().mapSettings().destinationCrs().authid()
                                canvas_crs = QgsCoordinateReferenceSystem(canvas_auth_id)
                                ellipsoid_crs = QgsCoordinateReferenceSystem(4326)
                                trans_context = QgsCoordinateTransformContext()
                                trans_context.calculateDatumTransforms(canvas_crs, ellipsoid_crs)
                                d.setSourceCrs(canvas_crs, trans_context)
                                dist = d.measureLine(closest_vertex[0], closest_vertex_test[0])  # distance in metres
                                dist_list.append(dist)
                                attr_list.append(test_feature.attribute(self.in_mpa_field))

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
                        ws.write(row, 2, dist_dict[item][1] / 1000)  # conversion from metres to kilometres
                        row += 1
                
                # loop through polygon layers
                name_list = []
                for polyName in self.check_poly_dict.keys():
                    # add a new worksheet to workbook, length limit is 31 characters
                    short_name = polyName[:30]
                    if short_name in name_list:
                        short_name = short_name[:29] + "1"
                    name_list.append(short_name)
                    ws = wb.add_sheet(short_name)
                    row = 1
                    # define green and red styles
                    green_style = 'pattern: pattern solid, pattern_fore_colour lime, pattern_back_colour lime'
                    red_style = 'pattern: pattern solid, pattern_fore_colour rose, pattern_back_colour rose'
                    # get information from the dictionary
                    layer = self.check_poly_dict[polyName]['layer']
                    field = self.check_poly_dict[polyName]['field']
                    coverage_target = self.check_poly_dict[polyName]['coverageTarget']
                    repl_target = self.check_poly_dict[polyName]['replTarget']
                    # write header
                    header_cells = [polyName + " " + field.name(),
                                    "Coverage" + " target=" + "{0:.0f}%".format(coverage_target),
                                    "Replication" + ' target=' + str(repl_target)]
                    for i in range(len(header_cells)):
                        ws.write(0, i, header_cells[i])
                    # create list of unique IDs for polygons
                    attr_index = layer.fields().lookupField(field.name())
                    attr_list = layer.uniqueValues(attr_index)
                    attr_list = list(attr_list)
                    attr_list.sort()
                    # create dictionary with entry for each polygon with values of area intersecting with each MPA
                    mpa_area_per_poly = intersect_area(layer, field.name(), self.in_map_layer, self.in_mpa_field)
                    # print report 
                    for uniqueID in attr_list:
                        sum_area = sum(mpa_area_per_poly[uniqueID].values())
                        mpa_count = str(len([PA for PA in mpa_area_per_poly[uniqueID]]))
                        print_list = [uniqueID, sum_area, mpa_count]
                        for attribute in print_list:
                            if attribute is None:
                                attribute = "NULL"
                            if attribute == uniqueID:
                                try:
                                    attribute = int(uniqueID)
                                except ValueError:
                                    pass
                                ws.write(row, 0, attribute)
                            elif attribute == sum_area:
                                attribute = float(attribute)
                                if attribute >= coverage_target/100.0:
                                    style_string = green_style
                                else:
                                    style_string = red_style
                                style = xlwt.easyxf(style_string, num_format_str='0%')
                                ws.write(row, 1, attribute, style)
                            elif attribute == mpa_count:
                                attribute = int(attribute)
                                if attribute >= repl_target:
                                    style_string = green_style
                                else:
                                    style_string = red_style
                                style = xlwt.easyxf(style_string)
                                ws.write(row, 2, attribute, style)
                        row += 1
                    for i in range(len(header_cells)):
                        ws.col(i).width = (len(header_cells[i]) + 4) * 367
                wb.save(self.out_xls)
                if os.path.exists(self.out_xls):
                    try:
                        from os import startfile  # windows only
                        os.startfile(self.out_xls)
                    except ImportError:
                        import subprocess
                        subprocess.run(['xdg-open', self.out_xls])  # if not windows

            self.dlg_targets.outButton.clicked.disconnect(out_file)
