import os
import string
from qgis.PyQt.QtWidgets import (
    QAction, QDialog, QVBoxLayout, QLabel, QComboBox, QLineEdit, QHBoxLayout,
    QCheckBox, QPushButton, QMessageBox, QProgressDialog, QApplication
)
from qgis.PyQt.QtGui import QIcon, QFont, QColor
from PyQt5.QtCore import QVariant, QThread, pyqtSignal
from qgis.core import (
    QgsProject, QgsCoordinateReferenceSystem, QgsRectangle,
    QgsVectorLayer, QgsFeature, QgsGeometry, QgsField,
    QgsFillSymbol, QgsPalLayerSettings, QgsTextFormat,
    QgsTextBufferSettings, QgsVectorLayerSimpleLabeling,
    QgsCoordinateTransform, QgsCoordinateTransformContext
)

# --- Worker Thread for background grid creation ---
class GridGeneratorThread(QThread):
    progressChanged = pyqtSignal(int)
    finished = pyqtSignal(list)

    def __init__(self, transformed_features, to_original, grid_width, grid_height, xmin, xmax, ymin, ymax):
        super().__init__()
        self.transformed_features = transformed_features
        self.to_original = to_original
        self.grid_width = grid_width
        self.grid_height = grid_height
        self.xmin = xmin
        self.xmax = xmax
        self.ymin = ymin
        self.ymax = ymax

    def run(self):
        features = []
        row, y = 1, self.ymin
        total_rows = int((self.ymax - self.ymin) / self.grid_height) + 1
        processed_rows = 0
        while y < self.ymax:
            x, col = self.xmin, 1
            while x < self.xmax:
                rect = QgsRectangle(x, y, x + self.grid_width, y + self.grid_height)
                geom_rect = QgsGeometry.fromRect(rect)
                if any(geom.intersects(geom_rect) for geom in self.transformed_features):
                    feat = QgsFeature()
                    feat.setGeometry(QgsGeometry.fromRect(self.to_original.transformBoundingBox(rect)))
                    features.append((feat, feat.geometry().centroid().asPoint(), row, col))
                x += self.grid_width
                col += 1
            y += self.grid_height
            row += 1
            processed_rows += 1
            percent = int(100 * processed_rows / total_rows)
            self.progressChanged.emit(percent)
        self.finished.emit(features)

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
        self.scale_options = [
            "1:500", "1:750", "1:1000", "1:1500", "1:2000", "1:2500", "1:3000", "1:4000", "1:5000",
            "1:7500", "1:10000", "1:15000", "1:20000", "1:25000"
        ]
        self.scale_combo.addItems(self.scale_options)
        layout.addWidget(self.scale_combo)

        # --- Box and checkbox for custom scale input ---
        scale_input_layout = QHBoxLayout()
        self.custom_scale_checkbox = QCheckBox("Oder Maßstab eingeben:")
        self.scale_input = QLineEdit()
        self.scale_input.setPlaceholderText("z.B. 300")
        self.scale_input.setEnabled(False)
        scale_label = QLabel("1:")
        scale_input_layout.addWidget(self.custom_scale_checkbox)
        scale_input_layout.addWidget(scale_label)
        scale_input_layout.addWidget(self.scale_input)
        layout.addLayout(scale_input_layout)

        self.custom_scale_checkbox.stateChanged.connect(self.toggle_scale_mode)

        layout.addWidget(QLabel("Layout-Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["quer", "hoch"])
        layout.addWidget(self.format_combo)

        layout.addWidget(QLabel("Papiergröße:"))
        self.paper_combo = QComboBox()
        self.paper_combo.addItems([
            "A6", "A5", "A4", "A3", "A2", "A1", "A0",
            "B6", "B5", "B4", "B3", "B2", "B1", "B0",
            "Letter", "Legal",
            "ANSI A", "ANSI B", "ANSI C", "ANSI D", "ANSI E",
            "Arch A", "Arch B", "Arch C", "Arch D", "Arch E", "Arch E1", "Arch E2", "Arch E3"
        ])
        layout.addWidget(self.paper_combo)

        # manually adding Breite/Höhe (mm)
        self.manual_size_checkbox = QCheckBox("Kartenausschnitt manuell festlegen (Breite/Höhe in mm):")
        layout.addWidget(self.manual_size_checkbox)

        manual_size_layout = QHBoxLayout()
        self.manual_width = QLineEdit()
        self.manual_width.setPlaceholderText("Breite (mm)")
        self.manual_height = QLineEdit()
        self.manual_height.setPlaceholderText("Höhe (mm)")
        manual_size_layout.addWidget(QLabel("Breite:"))
        manual_size_layout.addWidget(self.manual_width)
        manual_size_layout.addWidget(QLabel("Höhe:"))
        manual_size_layout.addWidget(self.manual_height)
        layout.addLayout(manual_size_layout)

        self.manual_width.setEnabled(False)
        self.manual_height.setEnabled(False)
        self.manual_size_checkbox.stateChanged.connect(
            lambda: [
                self.manual_width.setEnabled(self.manual_size_checkbox.isChecked()),
                self.manual_height.setEnabled(self.manual_size_checkbox.isChecked()),
                self.paper_combo.setEnabled(not self.manual_size_checkbox.isChecked())
            ]
        )

        run_button = QPushButton("Gitter erstellen")
        run_button.clicked.connect(lambda: self.generate_grid(dialog))
        layout.addWidget(run_button)

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        dialog.setLayout(layout)
        dialog.setMinimumWidth(350)
        dialog.exec_()

    def toggle_scale_mode(self):
        is_custom = self.custom_scale_checkbox.isChecked()
        self.scale_combo.setEnabled(not is_custom)
        self.scale_input.setEnabled(is_custom)

    # --- For each paper type, (height, width) in mm ---
    paper_sizes_mm = {
        "A6": (105, 148),
        "A5": (148, 210),
        "A4": (210, 297),
        "A3": (297, 420),
        "A2": (420, 594),
        "A1": (594, 841),
        "A0": (841, 1189),
        "B6": (125, 176),
        "B5": (176, 250),
        "B4": (250, 353),
        "B3": (353, 500),
        "B2": (500, 707),
        "B1": (707, 1000),
        "B0": (1000, 1414),
        "Letter": (216, 279),
        "Legal": (216, 356),
        "ANSI A": (216, 279),
        "ANSI B": (279, 432),
        "ANSI C": (432, 559),
        "ANSI D": (559, 864),
        "ANSI E": (864, 1118),
        "Arch A": (229, 305),
        "Arch B": (305, 457),
        "Arch C": (457, 610),
        "Arch D": (610, 914),
        "Arch E": (914, 1219),
        "Arch E1": (762, 1067),
        "Arch E2": (660, 965),
        "Arch E3": (686, 991)
    }

    def get_default_projected_crs(self, layer):
        extent = layer.extent()
        if extent.isEmpty():
            return QgsCoordinateReferenceSystem("EPSG:25832")
        # Calculate UTM zone
        center_x = (extent.xMinimum() + extent.xMaximum()) / 2.0
        center_y = (extent.yMinimum() + extent.yMaximum()) / 2.0
        if layer.crs().isGeographic():
            # Determine UTM zone by longitude
            zone = int((center_x + 180) / 6) + 1
            epsg_code = 32600 + zone if center_y >= 0 else 32700 + zone
            return QgsCoordinateReferenceSystem(f"EPSG:{epsg_code}")
        else:
            return layer.crs()

    def get_column_label(self, index):
        result = ""
        index -= 1
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    # --- NEW: Add grid features after thread completion ---
    def add_grid_features(self, features, grid_layer, provider):
        sorted_feats = sorted(features, key=lambda x: (-x[1].y(), x[1].x()))
        for i, (feat, _, row, col) in enumerate(sorted_feats):
            label = f"{self.get_column_label(col)}{row}"
            feat.setFields(grid_layer.fields())
            feat.setAttribute("grid", label)
            feat.setAttribute("serial", i + 1)
        provider.addFeatures([f[0] for f in sorted_feats])

    def generate_grid(self, dialog):
        layer_name = self.layer_combo.currentText()
        orientation = self.format_combo.currentText()
        paper_size = self.paper_combo.currentText()
        # Scale selection
        if self.custom_scale_checkbox.isChecked():
            try:
                user_scale = int(self.scale_input.text().strip())
                if user_scale <= 0:
                    raise ValueError()
            except Exception:
                QMessageBox.warning(None, "Fehler", "Bitte eine gültige Maßstabzahl eingeben!")
                return
            scale = user_scale
        else:
            scale_label = self.scale_combo.currentText().replace("1:", "")
            scale = int(scale_label)

        # Grid size (Width/Height): If there is manual entry, it will be calculated first, otherwise it will be calculated automatically.
        if self.manual_size_checkbox.isChecked():
            try:
                grid_width_mm = float(self.manual_width.text().replace(",", "."))
                grid_height_mm = float(self.manual_height.text().replace(",", "."))
                if grid_width_mm <= 0 or grid_height_mm <= 0:
                    raise ValueError()
            except Exception:
                QMessageBox.warning(None, "Fehler", "Bitte gültige Werte für Breite und Höhe eingeben!")
                return
        else:
            height_mm, width_mm = self.paper_sizes_mm[paper_size]
            if orientation == "quer":
                grid_width_mm = width_mm
                grid_height_mm = height_mm
            else:
                grid_width_mm = height_mm
                grid_height_mm = width_mm

        # Converts mm to meters and multiplies by scale:
        grid_width = grid_width_mm / 1000 * scale  # meters
        grid_height = grid_height_mm / 1000 * scale

        layers = QgsProject.instance().mapLayersByName(layer_name)
        if not layers or layer_name == "Kein Vektor-Layer gefunden":
            QMessageBox.warning(None, "Fehler", "Der ausgewählte Layer wurde nicht gefunden oder ist ungültig.")
            dialog.close()
            return
        layer = layers[0]
        crs = layer.crs()
        transform_context = QgsProject.instance().transformContext()
        #  First select a suitable metric CRS for the grid
        target_crs = self.get_default_projected_crs(layer)
        to_target = QgsCoordinateTransform(crs, target_crs, transform_context)
        to_original = QgsCoordinateTransform(target_crs, crs, transform_context)

        #  Transform the layer to the metric CRS and create the grid here
        total = layer.featureCount()
        progress = QProgressDialog("Verarbeitung läuft...", None, 0, 100)
        progress.setWindowTitle("Bitte warten")
        progress.setWindowModality(True)
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()
        transformed_features = []
        for i, feature in enumerate(layer.getFeatures()):
            geom = QgsGeometry(feature.geometry())
            geom.transform(to_target)
            if not geom.isEmpty():
                transformed_features.append(geom)
            QApplication.processEvents()
        if not transformed_features:
            progress.close()
            QMessageBox.information(None, "Hinweis", "Keine gültigen Geometrien im Layer.")
            dialog.close()
            return

        bounds = QgsGeometry.unaryUnion(transformed_features).boundingBox()
        offset = 10
        xmin, xmax = bounds.xMinimum() - offset, bounds.xMaximum() + offset
        ymin, ymax = bounds.yMinimum() - offset, bounds.yMaximum() + offset

        if self.manual_size_checkbox.isChecked():
            try:
                grid_width_mm = float(self.manual_width.text().replace(",", "."))
                grid_height_mm = float(self.manual_height.text().replace(",", "."))
                if grid_width_mm <= 0 or grid_height_mm <= 0:
                    raise ValueError()
                size_string = f"{int(grid_width_mm)}x{int(grid_height_mm)}mm"
            except Exception:
                QMessageBox.warning(None, "Fehler", "Bitte gültige Werte für Breite und Höhe eingeben!")
                progress.close()
                return
        else:
            height_mm, width_mm = self.paper_sizes_mm[paper_size]
            if orientation == "quer":
                grid_width_mm = width_mm
                grid_height_mm = height_mm
            else:
                grid_width_mm = height_mm
                grid_height_mm = width_mm
            size_string = paper_size

        layer_base_name = layer.name().replace(" ", "_").replace(":", "_")
        base_name = f"Gitter_1:{scale}_{orientation}_{size_string}_{layer_base_name}"
        existing_names = [l.name() for l in QgsProject.instance().mapLayers().values()]
        counter = 1
        grid_layer_name = f"{base_name}_{counter:02d}"
        while grid_layer_name in existing_names:
            counter += 1
            grid_layer_name = f"{base_name}_{counter:02d}"

        #  Create grids in the metric CRS (for most accurate results)
        grid_layer = QgsVectorLayer(f"Polygon?crs={target_crs.authid()}", grid_layer_name, "memory")
        provider = grid_layer.dataProvider()
        provider.addAttributes([QgsField("grid", QVariant.String), QgsField("serial", QVariant.Int)])
        grid_layer.updateFields()

        # --- Start grid generator thread ---
        self.worker = GridGeneratorThread(
            transformed_features, to_original, grid_width, grid_height, xmin, xmax, ymin, ymax
        )

        def on_progress(val):
            progress.setValue(val)
            QApplication.processEvents()

        def on_finished(features):
            progress.close()
            self.add_grid_features(features, grid_layer, provider)
            grid_layer.updateExtents()
            if crs.authid() != target_crs.authid():
                # Optional: If needed, you can transform the grid features here (for final CRS)
                pass
            QgsProject.instance().addMapLayer(grid_layer)
            # Visual settings and labeling
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
            QMessageBox.information(None, "Fertig", f"{len(features)} Gitterzellen wurden erstellt.")
            dialog.close()

        self.worker.progressChanged.connect(on_progress)
        self.worker.finished.connect(on_finished)
        self.worker.start()
