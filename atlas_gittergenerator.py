import os
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
        self.scale_options = ["1:500", "1:750", "1:1000", "1:1500", "1:2000", "1:3000", "1:4000", "1:5000", "1:10000", "1:15000" , "1:20000", "1:25000"]
        self.scale_combo.addItems(self.scale_options)
        layout.addWidget(self.scale_combo)

        layout.addWidget(QLabel("Layout-Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["quer", "hoch"])
        layout.addWidget(self.format_combo)

        layout.addWidget(QLabel("Papiergröße:"))
        self.paper_combo = QComboBox()
        self.paper_combo.addItems(["A4", "A3"])
        layout.addWidget(self.paper_combo)

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

        if extent.isEmpty():
            QMessageBox.warning(None, "Fehler", "Layer hat keine gültige Geometrie oder ist leer.")
            return QgsCoordinateReferenceSystem("EPSG:25832")

        center_x = (extent.xMinimum() + extent.xMaximum()) / 2.0
        center_y = (extent.yMinimum() + extent.yMaximum()) / 2.0

        if layer.crs().isGeographic():
            try:
                zone = int((center_x + 180) / 6) + 1
                epsg_code = 32600 + zone if center_y >= 0 else 32700 + zone
                return QgsCoordinateReferenceSystem(f"EPSG:{epsg_code}")
            except ValueError:
                QMessageBox.warning(None, "Fehler", "Ungültiges Koordinatenzentrum erkannt.")
                return QgsCoordinateReferenceSystem("EPSG:25832")
        else:
            return layer.crs()
            
    def get_column_label(self, index):
        result = ""
        index -= 1
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    grid_sizes = {
        "A4": {
            "1:500": {"hoch": (95.0, 138.5), "quer": (138.5, 95.0)},
            "1:750": {"hoch": (142.5, 207.8), "quer": (207.8, 142.5)},
            "1:1000": {"hoch": (190.0, 277.0), "quer": (277.0, 190.0)},
            "1:1500": {"hoch": (285.0, 415.5), "quer": (415.5, 285.0)},
            "1:2000": {"hoch": (380.0, 554.0), "quer": (554.0, 380.0)},
            "1:3000": {"hoch": (570.0, 831.0), "quer": (831.0, 570.0)},
            "1:4000": {"hoch": (760.0, 1108.0), "quer": (1108.0, 760.0)},
            "1:5000": {"hoch": (950.0, 1385.0), "quer": (1385.0, 950.0)},
            "1:10000": {"hoch": (1900.0, 2770.0), "quer": (2770.0, 1900.0)},
            "1:15000": {"hoch": (2850.0, 4155.0), "quer": (4155.0, 2850.0)},
            "1:20000": {"hoch": (3800.0, 5540.0), "quer": (5540.0, 3800.0)},
            "1:25000": {"hoch": (4750.0, 6925.0), "quer": (6925.0, 4750.0)}
        },
        "A3": {
            "1:500": {"hoch": (138.5, 200.0), "quer": (200.0, 138.5)},
            "1:750": {"hoch": (207.8, 300.0), "quer": (300.0, 207.8)},
            "1:1000": {"hoch": (277.0, 400.0), "quer": (400.0, 277.0)},
            "1:1500": {"hoch": (415.5, 600.0), "quer": (600.0, 415.5)},
            "1:2000": {"hoch": (554.0, 800.0), "quer": (800.0, 554.0)},
            "1:3000": {"hoch": (831.0, 1200.0), "quer": (1200.0, 831.0)},
            "1:4000": {"hoch": (1108.0, 1600.0), "quer": (1600.0, 1108.0)},
            "1:5000": {"hoch": (1385.0, 2000.0), "quer": (2000.0, 1385.0)},
            "1:10000": {"hoch": (2770.0, 4000.0), "quer": (4000.0, 2770.0)},
            "1:15000": {"hoch": (4155.0, 6000.0), "quer": (6000.0, 4155.0)},
            "1:20000": {"hoch": (5540.0, 8000.0), "quer": (8000.0, 5540.0)},
            "1:25000": {"hoch": (6925.0, 10000.0), "quer": (10000.0, 6925.0)}
        }
    }

    def generate_grid(self, dialog):
        layer_name = self.layer_combo.currentText()
        scale_label = self.scale_combo.currentText()
        orientation = self.format_combo.currentText()
        paper_size = self.paper_combo.currentText()

        grid_width, grid_height = self.grid_sizes[paper_size][scale_label][orientation]

        layers = QgsProject.instance().mapLayersByName(layer_name)
        if not layers or layer_name == "Kein Vektor-Layer gefunden":
            QMessageBox.warning(None, "Fehler", "Der ausgewählte Layer wurde nicht gefunden oder ist ungültig.")
            dialog.close()
            return
        layer = layers[0]

        crs = layer.crs()
        transform_context = QgsProject.instance().transformContext()
        target_crs = self.get_default_projected_crs(layer)

        to_target = QgsCoordinateTransform(crs, target_crs, transform_context)
        to_original = QgsCoordinateTransform(target_crs, crs, transform_context)

        total = layer.featureCount()
        progress = QProgressDialog("Verarbeitung läuft...", None, 0, total)
        progress.setWindowTitle("Bitte warten")
        progress.setWindowModality(True)
        progress.show()

        transformed_features = []
        spatial_index = QgsSpatialIndex()
        for i, feature in enumerate(layer.getFeatures()):
            geom = QgsGeometry(feature.geometry())
            geom.transform(to_target)
            if not geom.isEmpty():
                new_feat = QgsFeature()
                new_feat.setGeometry(geom)
                transformed_features.append(geom)
                spatial_index.insertFeature(new_feat)
            progress.setValue(i + 1)
            QApplication.processEvents()

        if not transformed_features:
            QMessageBox.information(None, "Hinweis", "Keine gültigen Geometrien im Layer.")
            dialog.close()
            progress.close()
            return

        bounds = QgsGeometry.unaryUnion(transformed_features).boundingBox()
        offset = 10
        xmin, xmax = bounds.xMinimum() - offset, bounds.xMaximum() + offset
        ymin, ymax = bounds.yMinimum() - offset, bounds.yMaximum() + offset

        # Dinamik benzersiz isimlendirme
        layer_base_name = layer.name().replace(" ", "_").replace(":", "_")
        scale_clean = scale_label.replace(":", "_")
        base_name = f"Gitter_{scale_clean}_{paper_size}_{orientation}_{layer_base_name}"
        existing_names = [l.name() for l in QgsProject.instance().mapLayers().values()]
        counter = 1
        grid_layer_name = f"{base_name}_{counter:02d}"
        while grid_layer_name in existing_names:
            counter += 1
            grid_layer_name = f"{base_name}_{counter:02d}"

        grid_layer = QgsVectorLayer(f"Polygon?crs={crs.authid()}", grid_layer_name, "memory")
        provider = grid_layer.dataProvider()
        provider.addAttributes([QgsField("grid", QVariant.String), QgsField("serial", QVariant.Int)])
        grid_layer.updateFields()

        features = []
        row, y = 1, ymin
        while y < ymax:
            x, col = xmin, 1
            while x < xmax:
                rect = QgsRectangle(x, y, x + grid_width, y + grid_height)
                geom_rect = QgsGeometry.fromRect(rect)
                if any(geom.intersects(geom_rect) for geom in transformed_features):
                    feat = QgsFeature()
                    feat.setFields(grid_layer.fields())
                    feat.setGeometry(QgsGeometry.fromRect(to_original.transformBoundingBox(rect)))
                    label = f"{self.get_column_label(col)}{row}"
                    feat.setAttribute("grid", label)
                    feat.setAttribute("serial", 0)
                    features.append((feat, feat.geometry().centroid().asPoint()))
                x += grid_width
                col += 1
            y += grid_height
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
