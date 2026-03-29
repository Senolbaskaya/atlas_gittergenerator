import os

from qgis.PyQt.QtWidgets import (
    QAction, QDialog, QVBoxLayout, QLabel, QComboBox, QLineEdit, QHBoxLayout,
    QCheckBox, QPushButton, QMessageBox, QProgressDialog, QApplication
)
from qgis.PyQt.QtGui import QIcon, QFont, QColor
from PyQt5.QtCore import QVariant, QThread, pyqtSignal, QSettings, Qt

from qgis.core import (
    QgsProject,
    QgsCoordinateReferenceSystem,
    QgsRectangle,
    QgsVectorLayer,
    QgsFeature,
    QgsGeometry,
    QgsField,
    QgsFillSymbol,
    QgsPalLayerSettings,
    QgsTextFormat,
    QgsTextBufferSettings,
    QgsVectorLayerSimpleLabeling,
    QgsCoordinateTransform
)


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

    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.worker = None

    def tr(self, en, de):
        lang = QSettings().value("locale/userLocale", "en")
        if str(lang)[:2].lower() == "de":
            return de
        return en

    def initGui(self):
        icon_path = os.path.join(self.plugin_dir, "icon.png")
        self.action = QAction(
            QIcon(icon_path),
            self.tr("Atlas Grid Generator", "Atlas-Gittergenerator"),
            self.iface.mainWindow()
        )
        self.action.setToolTip(
            self.tr("Regular grid generator for atlas creation", "Regelmäßiges Gitter zur Atlas-Erstellung")
        )
        self.action.triggered.connect(self.show_dialog)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu(
            self.tr("Atlas Grid Generator", "Atlas-Gittergenerator"),
            self.action
        )

    def unload(self):
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu(
            self.tr("Atlas Grid Generator", "Atlas-Gittergenerator"),
            self.action
        )

    def show_dialog(self):
        dialog = QDialog()
        dialog.setWindowTitle(self.tr("Atlas Grid Generator", "Atlas-Gittergenerator"))
        layout = QVBoxLayout()

        layout.addWidget(QLabel(self.tr("Select input layer:", "Inhalt Layer wählen:")))
        self.layer_combo = QComboBox()
        vector_layers_found = False

        for layer in QgsProject.instance().mapLayers().values():
            if isinstance(layer, QgsVectorLayer):
                self.layer_combo.addItem(layer.name(), layer.name())
                vector_layers_found = True

        if not vector_layers_found:
            self.layer_combo.addItem(
                self.tr("No vector layer found", "Kein Vektor-Layer gefunden"),
                "NO_VECTOR_LAYER"
            )
            self.layer_combo.setEnabled(False)

        layout.addWidget(self.layer_combo)

        self.selected_only_checkbox = QCheckBox(
            self.tr(
                "Create grid only for selected features",
                "Nur für ausgewählte Objekte ein Gitter erstellen"
            )
        )
        layout.addWidget(self.selected_only_checkbox)

        layout.addWidget(QLabel(self.tr("Select scale:", "Maßstab wählen:")))
        self.scale_combo = QComboBox()
        self.scale_options = [
            "1:500", "1:750", "1:1000", "1:1500", "1:2000", "1:2500", "1:3000", "1:4000", "1:5000",
            "1:7500", "1:10000", "1:15000", "1:20000", "1:25000"
        ]
        for scale in self.scale_options:
            self.scale_combo.addItem(scale, scale)
        layout.addWidget(self.scale_combo)

        scale_input_layout = QHBoxLayout()
        self.custom_scale_checkbox = QCheckBox(
            self.tr("Or enter scale manually:", "Oder Maßstab eingeben:")
        )
        self.scale_input = QLineEdit()
        self.scale_input.setPlaceholderText(self.tr("e.g. 3000", "z.B. 3000"))
        self.scale_input.setEnabled(False)
        scale_label = QLabel("1:")
        scale_input_layout.addWidget(self.custom_scale_checkbox)
        scale_input_layout.addWidget(scale_label)
        scale_input_layout.addWidget(self.scale_input)
        layout.addLayout(scale_input_layout)

        self.custom_scale_checkbox.stateChanged.connect(self.toggle_scale_mode)

        layout.addWidget(QLabel(self.tr("Layout format:", "Layout-Format:")))
        self.format_combo = QComboBox()
        self.format_combo.addItem(self.tr("Landscape", "Querformat"), "landscape")
        self.format_combo.addItem(self.tr("Portrait", "Hochformat"), "portrait")
        layout.addWidget(self.format_combo)

        layout.addWidget(QLabel(self.tr("Paper size:", "Papiergröße:")))
        self.paper_combo = QComboBox()
        paper_items = [
            "A6", "A5", "A4", "A3", "A2", "A1", "A0",
            "B6", "B5", "B4", "B3", "B2", "B1", "B0",
            "Letter", "Legal",
            "ANSI A", "ANSI B", "ANSI C", "ANSI D", "ANSI E",
            "Arch A", "Arch B", "Arch C", "Arch D", "Arch E", "Arch E1", "Arch E2", "Arch E3"
        ]
        for paper in paper_items:
            self.paper_combo.addItem(paper, paper)
        layout.addWidget(self.paper_combo)

        self.manual_size_checkbox = QCheckBox(
            self.tr(
                "Set map extent manually (width/height in mm):",
                "Kartenausschnitt manuell festlegen (Breite/Höhe in mm):"
            )
        )
        layout.addWidget(self.manual_size_checkbox)

        manual_size_layout = QHBoxLayout()
        self.manual_width = QLineEdit()
        self.manual_width.setPlaceholderText(self.tr("Width (mm)", "Breite (mm)"))
        self.manual_height = QLineEdit()
        self.manual_height.setPlaceholderText(self.tr("Height (mm)", "Höhe (mm)"))
        manual_size_layout.addWidget(QLabel(self.tr("Width:", "Breite:")))
        manual_size_layout.addWidget(self.manual_width)
        manual_size_layout.addWidget(QLabel(self.tr("Height:", "Höhe:")))
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

        run_button = QPushButton(self.tr("Create grid", "Gitter erstellen"))
        run_button.clicked.connect(lambda: self.generate_grid(dialog))
        layout.addWidget(run_button)

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        dialog.setLayout(layout)
        dialog.setMinimumWidth(380)
        dialog.exec_()

    def toggle_scale_mode(self):
        is_custom = self.custom_scale_checkbox.isChecked()
        self.scale_combo.setEnabled(not is_custom)
        self.scale_input.setEnabled(is_custom)

    def get_default_projected_crs(self, layer):
        extent = layer.extent()
        if extent.isEmpty():
            return QgsCoordinateReferenceSystem("EPSG:25832")

        center_x = (extent.xMinimum() + extent.xMaximum()) / 2.0
        center_y = (extent.yMinimum() + extent.yMaximum()) / 2.0

        if layer.crs().isGeographic():
            zone = int((center_x + 180) / 6) + 1
            epsg_code = 32600 + zone if center_y >= 0 else 32700 + zone
            return QgsCoordinateReferenceSystem("EPSG:{0}".format(epsg_code))
        return layer.crs()

    def get_column_label(self, index):
        result = ""
        index -= 1
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    def add_grid_features(self, features, grid_layer, provider):
        sorted_feats = sorted(features, key=lambda x: (-x[1].y(), x[1].x()))
        prepared_feats = []

        for i, (feat, _, row, col) in enumerate(sorted_feats):
            label = "{0}{1}".format(self.get_column_label(col), row)
            feat.setFields(grid_layer.fields())
            feat.setAttribute("grid", label)
            feat.setAttribute("serial", i + 1)
            prepared_feats.append(feat)

        provider.addFeatures(prepared_feats)

    def generate_grid(self, dialog):
        layer_name = self.layer_combo.currentData()
        orientation = self.format_combo.currentData()
        paper_size = self.paper_combo.currentData()
        selected_only = self.selected_only_checkbox.isChecked()

        if layer_name == "NO_VECTOR_LAYER":
            QMessageBox.warning(
                None,
                self.tr("Error", "Fehler"),
                self.tr(
                    "The selected layer was not found or is invalid.",
                    "Der ausgewählte Layer wurde nicht gefunden oder ist ungültig."
                )
            )
            dialog.close()
            return

        if self.custom_scale_checkbox.isChecked():
            try:
                user_scale = int(self.scale_input.text().strip())
                if user_scale <= 0:
                    raise ValueError()
            except Exception:
                QMessageBox.warning(
                    None,
                    self.tr("Error", "Fehler"),
                    self.tr("Please enter a valid scale value.", "Bitte eine gültige Maßstabzahl eingeben!")
                )
                return
            scale = user_scale
        else:
            scale_label = self.scale_combo.currentData().replace("1:", "")
            scale = int(scale_label)

        if self.manual_size_checkbox.isChecked():
            try:
                grid_width_mm = float(self.manual_width.text().replace(",", "."))
                grid_height_mm = float(self.manual_height.text().replace(",", "."))
                if grid_width_mm <= 0 or grid_height_mm <= 0:
                    raise ValueError()
            except Exception:
                QMessageBox.warning(
                    None,
                    self.tr("Error", "Fehler"),
                    self.tr(
                        "Please enter valid width and height values.",
                        "Bitte gültige Werte für Breite und Höhe eingeben!"
                    )
                )
                return
        else:
            height_mm, width_mm = self.paper_sizes_mm[paper_size]
            if orientation == "landscape":
                grid_width_mm = width_mm
                grid_height_mm = height_mm
            else:
                grid_width_mm = height_mm
                grid_height_mm = width_mm

        grid_width = grid_width_mm / 1000.0 * scale
        grid_height = grid_height_mm / 1000.0 * scale

        layers = QgsProject.instance().mapLayersByName(layer_name)
        if not layers:
            QMessageBox.warning(
                None,
                self.tr("Error", "Fehler"),
                self.tr(
                    "The selected layer was not found or is invalid.",
                    "Der ausgewählte Layer wurde nicht gefunden oder ist ungültig."
                )
            )
            dialog.close()
            return

        layer = layers[0]
        crs = layer.crs()
        transform_context = QgsProject.instance().transformContext()

        target_crs = self.get_default_projected_crs(layer)
        to_target = QgsCoordinateTransform(crs, target_crs, transform_context)
        to_original = QgsCoordinateTransform(target_crs, crs, transform_context)

        if selected_only:
            if layer.selectedFeatureCount() == 0:
                QMessageBox.warning(
                    None,
                    self.tr("Error", "Fehler"),
                    self.tr("No features are selected.", "Es sind keine Objekte ausgewählt.")
                )
                return
            source_features = layer.getSelectedFeatures()
        else:
            source_features = layer.getFeatures()

        progress = QProgressDialog(
            self.tr("Processing...", "Verarbeitung läuft..."),
            None,
            0,
            100
        )
        progress.setWindowTitle(self.tr("Please wait", "Bitte warten"))
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()

        transformed_features = []
        for feature in source_features:
            geom = QgsGeometry(feature.geometry())
            geom.transform(to_target)
            if not geom.isEmpty():
                transformed_features.append(geom)
            QApplication.processEvents()

        if not transformed_features:
            progress.close()
            QMessageBox.information(
                None,
                self.tr("Information", "Hinweis"),
                self.tr("No valid geometries found in the layer.", "Keine gültigen Geometrien im Layer.")
            )
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
                size_string = "{0}x{1}mm".format(int(grid_width_mm), int(grid_height_mm))
            except Exception:
                QMessageBox.warning(
                    None,
                    self.tr("Error", "Fehler"),
                    self.tr(
                        "Please enter valid width and height values.",
                        "Bitte gültige Werte für Breite und Höhe eingeben!"
                    )
                )
                progress.close()
                return
        else:
            height_mm, width_mm = self.paper_sizes_mm[paper_size]
            if orientation == "landscape":
                grid_width_mm = width_mm
                grid_height_mm = height_mm
            else:
                grid_width_mm = height_mm
                grid_height_mm = width_mm
            size_string = paper_size

        layer_base_name = layer.name().replace(" ", "_").replace(":", "_")
        source_mode = "selected" if selected_only else "layer"
        base_name = "Gitter_1:{0}_{1}_{2}_{3}_{4}".format(
            scale, orientation, size_string, source_mode, layer_base_name
        )

        existing_names = [l.name() for l in QgsProject.instance().mapLayers().values()]
        counter = 1
        grid_layer_name = "{0}_{1:02d}".format(base_name, counter)
        while grid_layer_name in existing_names:
            counter += 1
            grid_layer_name = "{0}_{1:02d}".format(base_name, counter)

        grid_layer = QgsVectorLayer("Polygon?crs={0}".format(target_crs.authid()), grid_layer_name, "memory")
        provider = grid_layer.dataProvider()
        provider.addAttributes([
            QgsField("grid", QVariant.String),
            QgsField("serial", QVariant.Int)
        ])
        grid_layer.updateFields()

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

            QgsProject.instance().addMapLayer(grid_layer)

            symbol = QgsFillSymbol.createSimple({
                "color": "102,255,230,100",
                "outline_color": "0,0,128",
                "outline_width": "0.6"
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

            if selected_only:
                msg = self.tr(
                    "{0} grid cells were created only for the selected features.".format(len(features)),
                    "{0} Gitterzellen wurden nur für die ausgewählten Objekte erstellt.".format(len(features))
                )
            else:
                msg = self.tr(
                    "{0} grid cells were created.".format(len(features)),
                    "{0} Gitterzellen wurden erstellt.".format(len(features))
                )

            QMessageBox.information(
                None,
                self.tr("Done", "Fertig"),
                msg
            )
            dialog.close()

        self.worker.progressChanged.connect(on_progress)
        self.worker.finished.connect(on_finished)
        self.worker.start()