# Optimize
from qgis.core import *
from qgis.PyQt import QtGui
from qgis.utils import iface
from qgis.core import QgsProject, QgsCoordinateReferenceSystem
from qgis.PyQt.QtCore import QSize
from PyQt5.QtCore import QRectF
from qgis.PyQt.QtGui import QIcon, QColor, QKeySequence, QFont
from qgis.core import QgsLayoutItemShape
from qgis.PyQt.QtGui import QColor, QBrush
import os.path

# Main Input
main_layer_name = 'Kawasan Hutan'
layers_to_activate = ['osm_transport', 'batas_administrasi', 'Sungai','osm_roads','osm_water_a','Kawasan Hutan','ESRI World Topo','BIG_KECAMATAN_INDO','BIG_PROVINSI_INDO','BIG_KABKOT_INDO']

# Read Axisting Layer
Layers = QgsProject.instance().mapLayersByName(main_layer_name)
main_layer = Layers[0]

# Turn on All Layers to Avoid an Error
all_layers = QgsProject.instance().mapLayers().values()
# Iterate through the layers and set visibility to True
for layer in all_layers:
    QgsProject.instance().layerTreeRoot().findLayer(layer.id()).setItemVisibilityChecked(True)
    
def activate_layer(layers_to_activate):
    
    # Run to Delete the Unnecessery Layers: Iterate through all layers in the project
    for layer_id, layer in QgsProject.instance().mapLayers().items():
        # Check if the layer name is in the list of layers to activate
        if layer.name() in layers_to_activate:
            # Set the layer visibility to true (activate)
            QgsProject.instance().layerTreeRoot().findLayer(layer_id).setItemVisibilityChecked(True)
        else:
            # Set the layer visibility to false (deactivate)
            QgsProject.instance().layerTreeRoot().findLayer(layer_id).setItemVisibilityChecked(False)
#activate_layer(layers_to_activate)

# Adding layout Page in layout manager
project = QgsProject.instance()
manager = project.layoutManager()
layoutName = 'Automated Map'
layouts_list = manager.printLayouts()
for layout in layouts_list:
    if layout.name() == layoutName:
        manager.removeLayout(layout)
### Add layout 
layout = QgsPrintLayout(project)
layout.initializeDefaults()
layout.setName(layoutName)
manager.addLayout(layout)

def main_frame_layout(layout):
    # Paper Box
    rectangle_shape = QgsLayoutItemShape(layout)
    rectangle_shape.setShapeType(QgsLayoutItemShape.Rectangle)
    rectangle_shape.attemptMove(QgsLayoutPoint(0.150, 0.150, QgsUnitTypes.LayoutMillimeters))
    rectangle_shape.attemptResize(QgsLayoutSize(296.700, 209.700, QgsUnitTypes.LayoutMillimeters))
    paperbox = layout.addLayoutItem(rectangle_shape)
    # Main Map box
    rectangle_shape = QgsLayoutItemShape(layout)
    rectangle_shape.setShapeType(QgsLayoutItemShape.Rectangle)
    rectangle_shape.attemptMove(QgsLayoutPoint(1.688, 1.708, QgsUnitTypes.LayoutMillimeters))
    rectangle_shape.attemptResize(QgsLayoutSize(225.868, 206.609, QgsUnitTypes.LayoutMillimeters))
    mapbox = layout.addLayoutItem(rectangle_shape)
    # Edge Information Box
    rectangle_shape = QgsLayoutItemShape(layout)
    rectangle_shape.setShapeType(QgsLayoutItemShape.Rectangle)
    rectangle_shape.attemptMove(QgsLayoutPoint(229.334, 1.708, QgsUnitTypes.LayoutMillimeters))
    rectangle_shape.attemptResize(QgsLayoutSize(66.018, 206.609, QgsUnitTypes.LayoutMillimeters))
    edgebox = layout.addLayoutItem(rectangle_shape)
    # Insert Box
    rectangle_shape = QgsLayoutItemShape(layout)
    rectangle_shape.setShapeType(QgsLayoutItemShape.Rectangle)
    rectangle_shape.attemptMove(QgsLayoutPoint(230.550, 148.410, QgsUnitTypes.LayoutMillimeters))
    rectangle_shape.attemptResize(QgsLayoutSize(63.559, 33.155, QgsUnitTypes.LayoutMillimeters))
    insertbox = layout.addLayoutItem(rectangle_shape)

    def add_polyline(start_pointX, start_pointY, end_pointX, end_pointY, layout):
        # Recipe from layout Polyline
        # Mainly borrowed from https://github.com/qgis/QGIS/blob/master/tests/src/python/test_qgslayoutpolyline.py
        point = QPolygonF()
        point.append(QPointF(start_pointX, start_pointY))
        point.append(QPointF(end_pointX, end_pointY))
        #polygon2.append(QPointF(250.0, 100.0))
        #polygon2.append(QPointF(10.0, 200.0))
        layoutItemPolyline = QgsLayoutItemPolyline(point, layout)
        return layout.addLayoutItem(layoutItemPolyline)
        
    # Line 1
    line1 = add_polyline(229.335, 34.74,295.35, 34.74, layout)
    # Line 2
    line2 = add_polyline(229.335, 140.697, 295.35, 140.697, layout)
    # Line 3
    line3 = add_polyline(229.335, 182.792, 295.35, 182.792, layout)


    return paperbox, mapbox, edgebox, insertbox, line1, line2, line3

paperbox, mapbox, edgebox, insertbox, line1, line2, line3 = main_frame_layout(layout)

layer_extent = main_layer.extent()
main_xmin = layer_extent.xMinimum()-(0.2187/100*layer_extent.xMinimum())
main_ymin = layer_extent.yMinimum()-(-1.7925/100*layer_extent.yMinimum())
main_xmax = layer_extent.xMaximum()-(-0.2395/100*layer_extent.xMaximum())
main_ymax = layer_extent.yMaximum()-(3.2613/100*layer_extent.yMaximum())

def add_main_map(layer, layout):
    # Create Map Items in The Layout
    map = QgsLayoutItemMap(layout)
    map.setRect(10,10,10,10)
    map.zoomToExtent(layer.extent())
    map.setCrs(QgsCoordinateReferenceSystem('EPSG:4326'))
    
    # Set map extent
    layer_extent = layer.extent()
    xmin = layer_extent.xMinimum()-(0.2187/100*layer_extent.xMinimum())
    ymin = layer_extent.yMinimum()-(-1.7925/100*layer_extent.yMinimum())
    xmax = layer_extent.xMaximum()-(-0.2395/100*layer_extent.xMaximum())
    ymax = layer_extent.yMaximum()-(3.2613/100*layer_extent.yMaximum())
    
    # map.setExtent(iface.mapCanvas().extent())
    map.zoomToExtent(QgsRectangle(QgsPointXY(xmin, ymin), QgsPointXY(xmax, ymax)))
    #layer_extent = layer.extent()
    #layer_extent.grow(0.1)
    #map.zoomToExtent(layer_extent)
    #map.setBackgroundColor(QColor(255,255,255, 0))
    layers_to_activate = ['osm_transport', 'batas_administrasi', 'Sungai','osm_roads','osm_water_a','Kawasan Hutan','ESRI World Topo','BIG_KECAMATAN_INDO','BIG_PROVINSI_INDO','BIG_KABKOT_INDO']
    
    layer_list = [layer for layer in iface.mapCanvas().layers() if layer.name() in layers_to_activate]
    #layer_list01 = iface.mapCanvas().layers()
    #layer01 = layer_list01[0]
    map.setLayers(layer_list)
    
    layout.addLayoutItem(map)
    map.setFrameEnabled(True)
    map.attemptMove(QgsLayoutPoint(7.704, 7.738, QgsUnitTypes.LayoutMillimeters))
    map.attemptResize(QgsLayoutSize(213.513, 194.523, QgsUnitTypes.LayoutMillimeters))
    map.storeCurrentLayerStyles()
    map.setKeepLayerSet(True)
    map.setKeepLayerStyles(True)
    return map
    
main_map = add_main_map(main_layer, layout)


def add_scalebar_mainmap(map,layout):
    # Add Scalebar
    scalebar = QgsLayoutItemScaleBar(layout)
    scalebar.setStyle('Single Box')
    scalebar.setUnits(QgsUnitTypes.DistanceKilometers)
    scalebar.setNumberOfSegments(2) # right
    scalebar.setNumberOfSegmentsLeft(0)
    scalebar.setHeight(2)
    #scalebar.setUnitsPerSegment(10)
    scalebar.setSegmentSizeMode(QgsScaleBarSettings.SegmentSizeMode(1)) # change mode of scalebar segmentation into fit segment Mode
    scalebar.setMinimumBarWidth(5)
    scalebar.setMaximumBarWidth(50)
    scalebar.setLabelBarSpace(2)
    scalebar.setLinkedMap(map)
    scalebar.setUnitLabel('km')
    scalebar.setFont(QFont('MS Shell Dlg 2', 10))
    scalebar.update()
    layout.addLayoutItem(scalebar)
    scalebar.attemptMove(QgsLayoutPoint(233.174, 22.256, QgsUnitTypes.LayoutMillimeters))

add_scalebar_mainmap(main_map,layout)

# Add North Arrow
def add_north_arrow(which_map, layout):
    # Add North Arrow
    north = QgsLayoutItemPicture(layout)
    #file_path = os.path.realpath(_file_)
    north.setPicturePath("E:/XJ/project/05. Carbon Project/Automated Desktop Study/northarrow/WindRose_LB_04_b2.svg")
    layout.addLayoutItem(north)
    #north.setReferencePoint(0)
    north.setBackgroundEnabled(False)

    if which_map == 'main':
        north.attemptMove(QgsLayoutPoint(252, 3.598, QgsUnitTypes.LayoutMillimeters))
        north.attemptResize(QgsLayoutSize(21,21, QgsUnitTypes.LayoutMillimeters))
    elif which_map == 'insert':
        north.attemptMove(QgsLayoutPoint(235.438, 166.010, QgsUnitTypes.LayoutMillimeters))
        north.attemptResize(QgsLayoutSize(8.027, 7.264, QgsUnitTypes.LayoutMillimeters))

    # Set the fill color to black
    black_color = QColor(Qt.black)
    north.setSvgFillColor(black_color)
    layout.addLayoutItem(north)

add_north_arrow(which_map= 'main', layout=layout)

def petunjukletakpeta(text_content, layout):
    # Add Text Label of Petunjuk Letak Peta
    plp = QgsLayoutItemLabel(layout)
    plp.setText(text_content)
    # Set font with bold attribute
    font = QFont('MS Shell Dlg 2', 8, False)
    font.setBold(True)
    plp.setFont(font)
    #set size of label item. this step seems a little pointless to me but it doesn't work without it
    plp.adjustSizeToText() 
    plp.setMarginX(3)
    plp.attemptMove(QgsLayoutPoint(230.550, 141.989, QgsUnitTypes.LayoutMillimeters))
    plp.attemptResize(QgsLayoutSize(63.559,5.152, QgsUnitTypes.LayoutMillimeters))
    plp.setBackgroundColor(QColor(31, 120, 180))
    plp.setBackgroundEnabled(True)
    plp.setFrameEnabled(True)
    plp.setFontColor(QColor(Qt.white))
    plp.setHAlign(Qt.AlignCenter)
    plp.setVAlign(Qt.AlignCenter)
    layout.addLayoutItem(plp)
    
petunjukletakpeta("Petunjuk Letak Peta",layout)
# add perunjuk letak peta map


def add_plp_map(layer, layout):
    
    layers_to_activate2 = ['Irmasulindo index','ESRI World Topo']
    #activate_layer(layers_to_activate2)
    
    # Create Map Items in The Layout
    map2 = QgsLayoutItemMap(layout)
    map2.setRect(10,10,10,10)
    map2.setCrs(QgsCoordinateReferenceSystem('EPSG:4326'))
    layer_extent = layer.extent()
    layer_extent.grow(7)#pass a sensible value depending on crs used and map scale
    map2.zoomToExtent(layer_extent)
    
    insert_layer = ['Irmasulindo index','ESRI World Topo']
    layer_list = [layer for layer in iface.mapCanvas().layers() if layer.name() in insert_layer]
    map2.setLayers(layer_list)
    #map.setBackgroundColor(QColor(255,255,255, 0))
    layout.addLayoutItem(map2)
    map2.setFrameEnabled(True)
    map2.attemptMove(QgsLayoutPoint(233.440,151.403, QgsUnitTypes.LayoutMillimeters))
    map2.attemptResize(QgsLayoutSize(57.512,27.123, QgsUnitTypes.LayoutMillimeters))
    map2.storeCurrentLayerStyles()
    map2.setKeepLayerSet(True)
    map2.setKeepLayerStyles(True)
    return map2
    
plp_map = add_plp_map(main_layer, layout)
add_north_arrow(which_map= 'insert', layout=layout)

def add_scalebar_mainmap(map,layout):
    # Add Scalebar
    scalebar = QgsLayoutItemScaleBar(layout)
    scalebar.setStyle('Line Ticks Up')
    scalebar.setUnits(QgsUnitTypes.DistanceKilometers)
    scalebar.setNumberOfSegments(2) # right
    scalebar.setNumberOfSegmentsLeft(0)
    scalebar.setHeight(1)
    #scalebar.setUnitsPerSegment(10)
    scalebar.setSegmentSizeMode(QgsScaleBarSettings.SegmentSizeMode(1)) # change mode of scalebar segmentation into fit segment Mode
    scalebar.setMinimumBarWidth(1)
    scalebar.setMaximumBarWidth(20)
    scalebar.setLabelBarSpace(1)
    scalebar.setLinkedMap(map)
    scalebar.setUnitLabel('km')
    scalebar.setFont(QFont('MS Shell Dlg 2', 5))
    scalebar.update()
    layout.addLayoutItem(scalebar)
    scalebar.attemptMove(QgsLayoutPoint(234.415, 171.910, QgsUnitTypes.LayoutMillimeters))
    scalebar.attemptResize(QgsLayoutSize(21.579, 6.088, QgsUnitTypes.LayoutMillimeters))

add_scalebar_mainmap(plp_map,layout)

def add_sumberdata(layout, text_content):
    submber_data = QgsLayoutItemLabel(layout)
    submber_data.setText(text_content)
    submber_data.setFont(QFont('MS Shell Dlg 2', 8))
    #set size of label item. this step seems a little pointless to me but it doesn't work without it
    submber_data.adjustSizeToText() 
    submber_data.setMarginX(3)
    submber_data.attemptMove(QgsLayoutPoint(229.5, 185, QgsUnitTypes.LayoutMillimeters))
    submber_data.attemptResize(QgsLayoutSize(66.000,21, QgsUnitTypes.LayoutMillimeters))
    layout.addLayoutItem(submber_data)
    
add_sumberdata(layout,
"""Sumber Peta 
-  AGB GlobBiomass, Santoro et al. 2018
-  Carbon Stock, UNEP-WCMC
-  BGB root shoot ratio, IPCC 2019 Refinement
-  Peta Wilayah Ina-Geoportal
-  OpenStreetMap""")


def add_grid (mymap, e, intvX, intvY, font_size, frame_thickness, frame_width, frame_distance, offsetX, offsetY):
    xmin = float(round(e.xMinimum(), 3))
    xmax = float(round(e.xMaximum(), 3))
    ymin = float(round(e.yMinimum(), 3))
    ymax = float(round(e.yMaximum(), 3))
    width = round(xmax-xmin,0)
    height = round(ymax-ymin,0)
    
    grid = QgsLayoutItemMapGrid('Grid 01', mymap)
    mymap.grids().addGrid(grid)
    grid.setCrs(QgsCoordinateReferenceSystem('EPSG:4326'))
    grid.setEnabled(True)
    grid.setStyle(3) # Grid Style
    grid.setFrameStyle(4) # Frame Style
    grid.setFramePenSize(frame_thickness) # Frame Thickness
    grid.setFrameWidth(frame_width) #Frame Size
    grid.setIntervalX(width/(width*intvX)) # Grid X Interval
    grid.setIntervalY(height/(height*intvY)) # Grid Y Interval
    grid.setOffsetX(offsetX)
    grid.setOffsetY(offsetY)
    grid.setAnnotationEnabled(True)
    grid.setAnnotationDirection(1,  QgsLayoutItemMapGrid.BorderSide(0))
    grid.setAnnotationDirection(1,  QgsLayoutItemMapGrid.BorderSide(1))
    grid.setAnnotationFrameDistance(frame_distance)
    grid.setAnnotationFormat(7)
    grid.setAnnotationFont(QFont('MS Shell Dlg 2', font_size))

# Main Map Grid 
add_grid(mymap = main_map, e = iface.mapCanvas().extent(), intvX = 6.5, intvY = 7.5, font_size = 8, frame_thickness = 1, frame_width = 1, frame_distance = 1, offsetX = 0, offsetY = 0)
# Insert Map Grid
add_grid(mymap = plp_map, e = iface.mapCanvas().extent(), intvX = 0.21, intvY = 0.27, font_size = 5, frame_thickness = 0.3, frame_width = 0.6, frame_distance = 0.3, offsetX = 0.5, offsetY = 2)

################
# Add Legend
# Gather visible layers in the project layer tree and create a list of map layer objects
# which are not checked, which we will subsequently remove from the legend model
tree_layers = project.layerTreeRoot().children()
checked_layers = [layer.name() for layer in tree_layers if layer.isVisible()]

# This adds a legend item to the Print Layout
legend = QgsLayoutItemLegend(layout)
legend.setTitle("Legend")  # Set the title
font = QFont('MS Shell Dlg 2', 12, False)
font.setBold(True)
legend.rstyle(QgsLegendStyle.Title).setFont(font)

layout.addLayoutItem(legend)
legend.attemptMove(QgsLayoutPoint(232.961, 35.291, QgsUnitTypes.LayoutMillimeters))
legend.attemptResize(QgsLayoutSize(66.000,21, QgsUnitTypes.LayoutMillimeters))
legend.setBackgroundEnabled(False)
# Get reference to the existing legend model and root group, then remove the unchecked layers
legend.setAutoUpdateModel(False)  # This line is important!!
# Without it, the unchecked layers will be removed not only from the layout legend
# but also from the table of contents panel and your project!!
# Layers to exclude from the legend

layers_to_include_in_legend = ['osm_transport', 'batas_administrasi', 'Sungai','osm_roads','Kawasan Hutan']
root = QgsLayerTree()
for lyr in iface.mapCanvas().layers():
        if lyr.name() in layers_to_include_in_legend:  
            root.addLayer(lyr)
        legend.model().setRootGroup(root)
        layout.addItem(legend)

#itemLlegend = composition.getComposerItemById('Legend')
#tree_layer = itemLlegend.modelV2().rootGroup().addLayer(layer)
#QgsLegendRenderer.setNodeLegendStyle(tree_layer, QgsComposerLegendStyle.Hidden)
#legend.adjustBoxSize()

# If only one legend within the layout
layoutItemLegend = [i for i in layout.items() if isinstance(i, QgsLayoutItemLegend)][0]
# Could also use the following if you defined an id 'mylegend' for the QgsLayoutItemLegend (commented here)
# layoutItemLegend = layout.itemById('mylegend')
model = layoutItemLegend.model()
tree_legend = model.rootGroup()

# Deactivate auto-refresh of the legend
layoutItemLegend.setAutoUpdateModel(False)
layoutItemLegend.updateLegend()

hidden_layername = ['osm_transport','batas_administrasi','osm_roads']
for layer in hidden_layername:
    main_layer_remove = QgsProject.instance().mapLayersByName(layer)[0]
    # get legend items 
    main_layer_tree_remove = legend.model().rootGroup().findLayer(main_layer_remove) # QgsLayerTreeLayer object
    QgsLegendRenderer.setNodeLegendStyle(main_layer_tree_remove, QgsLegendStyle.Hidden)
    
    
def remove_item_batasdesa():
    # define layer which legend has to be modified
    layer = QgsProject.instance().mapLayersByName('batas_administrasi')[0]

    # get legend items 
    model = legend.model()
    layer_tree = legend.model().rootGroup().findLayer(layer) # QgsLayerTreeLayer object
    
    # assuming we want to keep only first and third node
    symbols_to_remain = range(1, 4)
    # setting a new order and applying changes
    QgsMapLayerLegendUtils.setLegendNodeOrder(layer_tree, symbols_to_remain)
    model.refreshLayerLegend(layer_tree)

remove_item_batasdesa()



main_layer = QgsProject.instance().mapLayersByName('Kawasan Hutan')[0]
# get legend items 
main_model = legend.model()
main_layer_tree = legend.model().rootGroup().findLayer(main_layer) # QgsLayerTreeLayer object

QgsLegendRenderer.setNodeLegendStyle(main_layer_tree, QgsLegendStyle.Group)

font2 = QFont('MS Shell Dlg 2', 10, False)
font2.setBold(True)
legend.rstyle(QgsLegendStyle.Group).setFont(font2)

vector_layer = QgsProject.instance().mapLayersByName('Kawasan Hutan')[0]
legend_item = [i for i in layout.items() if isinstance(i, QgsLayoutItemLegend)][0]
lyr = legend_item.model().rootGroup().findLayer(vector_layer)  # switch from QgsVectorLayer to QgsLayerTreeLayer
lyr.setUseLayerName(False)  # Make the legend use a name different from the layer's
tree_layers = legend_item.model().rootGroup().children()  # get the legend's layer tree

layer_name = 'Kawasan Hutan'
for tr in tree_layers:
    if tr.name() == layer_name: # ensure you have the correct child node
        tr.setName("Kawasan Hutan")  # set the child node's new name
legend_item.updateLegend()  # Update the QgsLayerTreeModel

# Label Size
font3 = QFont('MS Shell Dlg 2', 10, False)
legend.rstyle(QgsLegendStyle.SymbolLabel).setFont(font3)

# Export to Image
output_path = 'E:/XJ/project/05. Carbon Project/Automated Desktop Study/map/output'
img_path = output_path+'/Kawasan Hutan.png'
exporter = QgsLayoutExporter(layout)
exporter.exportToImage(img_path,QgsLayoutExporter.ImageExportSettings())