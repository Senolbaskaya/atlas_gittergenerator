import datetime
import os
import string

from qgis.PyQt.QtWidgets import (
    QAction, QDialog, QVBoxLayout, QLabel,
    QComboBox, QPushButton, QMessageBox, QProgressDialog, QApplication
)
from qgis.PyQt.QtGui import QIcon, QFont, QColor
from PyQt5.QtCore import QVariant
from qgis.core import (
    QgsProject, QgsCoordinateReferenceSystem, QgsRectangle,
    QgsVectorLayer, QgsFeature, QgsGeometry, QgsField,
    QgsFillSymbol, QgsPalLayerSettings, QgsTextFormat,
    QgsTextBufferSettings, QgsVectorLayerSimpleLabeling,
    QgsCoordinateTransform, QgsCoordinateTransformContext,
    QgsSpatialIndex
)

class AtlasGitterGenerator:
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)

    def initGui(self):
        icon_path = os.path.join(self.plugin_dir, "icon.png")
        self.action = QAction(QIcon(icon_path), "Atlas-Gittergenerator", self.iface.mainWindow())
        self.action.setToolTip("Regelmäßiges Gitter zur Atlas-Erstellung")
        self.action.triggered.connect(self.show_dialog)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("Atlas-Gittergenerator", self.action)

    def unload(self):
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu("Atlas-Gittergenerator", self.action)

    def show_dialog(self):
        dialog = QDialog()
        dialog.setWindowTitle("Atlas-Gittergenerator")
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Inhalt Layer wählen:"))
        self.layer_combo = QComboBox()
        vector_layers_found = False
        for layer in QgsProject.instance().mapLayers().values():
            if isinstance(layer, QgsVectorLayer):
                self.layer_combo.addItem(layer.name())
                vector_layers_found = True

        if not vector_layers_found:
            self.layer_combo.addItem("Kein Vektor-Layer gefunden")
            self.layer_combo.setEnabled(False)

        layout.addWidget(self.layer_combo)

        layout.addWidget(QLabel("Maßstab wählen:"))
        self.scale_combo = QComboBox()
        self.scale_options = {"1:500": 500, "1:1000": 1000, "1:2000": 2000, "1:3000": 3000,
                              "1:5000": 5000, "1:10000": 10000, "1:25000": 25000}
        for scale in self.scale_options:
            self.scale_combo.addItem(scale)
        layout.addWidget(self.scale_combo)

        layout.addWidget(QLabel("Layout-Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["quer", "hoch"])
        layout.addWidget(self.format_combo)

        run_button = QPushButton("Gitter erstellen")
        run_button.clicked.connect(lambda: self.generate_grid(dialog))
        layout.addWidget(run_button)

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        dialog.setLayout(layout)
        dialog.setMinimumWidth(250)
        dialog.exec_()

    def get_default_projected_crs(self, layer):
        extent = layer.extent()
        center_x = (extent.xMinimum() + extent.xMaximum()) / 2.0
        center_y = (extent.yMinimum() + extent.yMaximum()) / 2.0

        if layer.crs().isGeographic():
            zone = int((center_x + 180) / 6) + 1
            epsg_code = 32600 + zone if center_y >= 0 else 32700 + zone
            return QgsCoordinateReferenceSystem(f"EPSG:{epsg_code}")
        else:
            return layer.crs()

    # Convert column index to A-Z label
    def get_column_label(self, index):
        result = ""
        index -= 1
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    def generate_grid(self, dialog):
        layer_name = self.layer_combo.currentText()
        scale_label = self.scale_combo.currentText()
        orientation = self.format_combo.currentText()

        layers = QgsProject.instance().mapLayersByName(layer_name)
        if not layers or layer_name == "Kein Vektor-Layer gefunden":
            QMessageBox.warning(None, "Fehler", "Der ausgewählte Layer wurde nicht gefunden oder ist ungültig.")
            dialog.close()
            return
        layer = layers[0]

        total = layer.featureCount()
        progress = QProgressDialog("Verarbeitung läuft...", None, 0, total)
        progress.setWindowTitle("Bitte warten")
        progress.setWindowModality(True)
        progress.setValue(0)
        progress.show()

        crs = layer.crs()
        transform_context = QgsProject.instance().transformContext()
        target_crs = self.get_default_projected_crs(layer)

        to_target = QgsCoordinateTransform(crs, target_crs, transform_context)
        to_original = QgsCoordinateTransform(target_crs, crs, transform_context)

        transformed_feats = []
        spatial_index = QgsSpatialIndex()
        for i, f in enumerate(layer.getFeatures()):
            geom = QgsGeometry(f.geometry())
            geom.transform(to_target)
            if not geom.isEmpty():
                new_feat = QgsFeature()
                new_feat.setId(i)
                new_feat.setGeometry(geom)
                transformed_feats.append(new_feat)
                spatial_index.insertFeature(new_feat)
            progress.setValue(i + 1)
            QApplication.processEvents()

        if not transformed_feats:
            QMessageBox.information(None, "Hinweis", "Keine gültigen Geometrien im Layer.")
            dialog.close()
            progress.close()
            return

        grid_sizes = {
            "1:500": {"hoch": (172, 140), "quer": (140, 172)},
            "1:1000": {"hoch": (344, 280), "quer": (280, 344)},
            "1:2000": {"hoch": (688, 560), "quer": (560, 688)},
            "1:3000": {"hoch": (1032, 840), "quer": (840, 1032)},
            "1:5000": {"hoch": (1720, 1400), "quer": (1400, 1720)},
            "1:10000": {"hoch": (3440, 2800), "quer": (2800, 3440)},
            "1:25000": {"hoch": (8600, 7000), "quer": (7000, 8600)}
        }
        height, width = grid_sizes[scale_label][orientation]

        layer_base_name = layer.name().replace(" ", "_").replace(":", "_")
        scale_clean = scale_label.replace(":", "_")
        base_name = f"Gitter_{scale_clean}_{orientation}_{layer_base_name}"

        existing_names = [l.name() for l in QgsProject.instance().mapLayers().values()]
        counter = 1
        grid_layer_name = f"{base_name}_{counter:02d}"

        while grid_layer_name in existing_names:
            counter += 1
            grid_layer_name = f"{base_name}_{counter:02d}"

        grid_layer = QgsVectorLayer("Polygon?crs=" + crs.authid(), grid_layer_name, "memory")
        provider = grid_layer.dataProvider()
        provider.addAttributes([QgsField("grid", QVariant.String), QgsField("serial", QVariant.Int)])
        grid_layer.updateFields()

        bounds = QgsGeometry.unaryUnion([f.geometry() for f in transformed_feats]).boundingBox()
        offset = 10
        xmin, xmax = bounds.xMinimum() - offset, bounds.xMaximum() + offset
        ymin, ymax = bounds.yMinimum() - offset, bounds.yMaximum() + offset

        features = []
        row, y = 1, ymin
        while y < ymax:
            x, col = xmin, 1
            while x < xmax:
                rect = QgsRectangle(x, y, x + width, y + height)
                ids = spatial_index.intersects(rect)
                if ids:
                    geom_rect = QgsGeometry.fromRect(rect)
                    intersecting = any(transformed_feats[idx].geometry().intersects(geom_rect) for idx in ids)
                    if intersecting:
                        feat = QgsFeature()
                        feat.setFields(grid_layer.fields())
                        geom_orig = QgsGeometry.fromRect(to_original.transformBoundingBox(rect))
                        feat.setGeometry(geom_orig)
                        label = f"{self.get_column_label(col)}{row}"
                        feat.setAttribute("grid", label)
                        feat.setAttribute("serial", 0)
                        features.append((feat, geom_orig.centroid().asPoint()))
                x += width
                col += 1
            y += height
            row += 1

        sorted_feats = sorted(features, key=lambda x: (-x[1].y(), x[1].x()))
        for i, (feat, _) in enumerate(sorted_feats):
            feat.setAttribute("serial", i + 1)

        provider.addFeatures([f[0] for f in sorted_feats])
        grid_layer.updateExtents()
        QgsProject.instance().addMapLayer(grid_layer)

        symbol = QgsFillSymbol.createSimple({
            'color': '102,255,230,100',
            'outline_color': '0,0,128',
            'outline_width': '0.6'
        })
        grid_layer.renderer().setSymbol(symbol)

        label_settings = QgsPalLayerSettings()
        text_format = QgsTextFormat()
        font = QFont("Arial", 10)
        font.setBold(True)
        text_format.setFont(font)
        text_format.setSize(10)

        buffer_settings = QgsTextBufferSettings()
        buffer_settings.setEnabled(True)
        buffer_settings.setSize(1)
        buffer_settings.setColor(QColor("white"))
        text_format.setBuffer(buffer_settings)

        label_settings.setFormat(text_format)
        label_settings.fieldName = "serial"
        label_settings.placement = QgsPalLayerSettings.AroundPoint
        label_settings.enabled = True

        grid_layer.setLabeling(QgsVectorLayerSimpleLabeling(label_settings))
        grid_layer.setLabelsEnabled(True)
        grid_layer.triggerRepaint()

        self.iface.layerTreeView().refreshLayerSymbology(grid_layer.id())
        progress.close()
        QMessageBox.information(None, "Fertig", f"{grid_layer.featureCount()} Gitterzellen wurden erstellt.")
        dialog.close()
