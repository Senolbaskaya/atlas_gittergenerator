import os
import math

from qgis.PyQt.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QComboBox, QLineEdit, QHBoxLayout,
    QCheckBox, QPushButton, QMessageBox, QProgressDialog, QApplication
)
from qgis.PyQt.QtGui import QAction, QIcon, QFont, QColor
from qgis.PyQt.QtCore import QThread, pyqtSignal, Qt, QMetaType, QSettings

from qgis.core import (
    Qgis,
    QgsProject,
    QgsCoordinateReferenceSystem,
    QgsCoordinateTransform,
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
    QgsPointXY
)


class GridGeneratorThread(QThread):
    progressChanged = pyqtSignal(int)
    finished = pyqtSignal(list)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, transformed_geometries, grid_width, grid_height, xmin, xmax, ymin, ymax):
        super().__init__()
        self.transformed_geometries = transformed_geometries
        self.grid_width = grid_width
        self.grid_height = grid_height
        self.xmin = xmin
        self.xmax = xmax
        self.ymin = ymin
        self.ymax = ymax
        self._cancel_requested = False

    def cancel(self):
        self._cancel_requested = True

    def run(self):
        try:
            if self.grid_width <= 0 or self.grid_height <= 0:
                self.failed.emit("Invalid grid size.")
                return

            total_rows = max(1, math.ceil((self.ymax - self.ymin) / self.grid_height))
            processed_rows = 0
            cells = []

            y = self.ymin
            row_from_bottom = 1

            while y < self.ymax:
                if self._cancel_requested:
                    self.cancelled.emit()
                    return

                x = self.xmin
                col = 1

                while x < self.xmax:
                    if self._cancel_requested:
                        self.cancelled.emit()
                        return

                    rect = QgsRectangle(x, y, x + self.grid_width, y + self.grid_height)
                    rect_geom = QgsGeometry.fromRect(rect)

                    if any(geom.intersects(rect_geom) for geom in self.transformed_geometries):
                        cx = x + (self.grid_width / 2.0)
                        cy = y + (self.grid_height / 2.0)
                        cells.append((
                            rect.xMinimum(),
                            rect.yMinimum(),
                            rect.xMaximum(),
                            rect.yMaximum(),
                            row_from_bottom,
                            col,
                            cx,
                            cy
                        ))

                    x += self.grid_width
                    col += 1

                y += self.grid_height
                row_from_bottom += 1
                processed_rows += 1

                percent = int((processed_rows / total_rows) * 100)
                self.progressChanged.emit(percent)

            self.finished.emit(cells)

        except Exception as e:
            self.failed.emit(str(e))


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
        return de if str(lang)[:2].lower() == "de" else en

    def initGui(self):
        icon_path = os.path.join(self.plugin_dir, "icon.png")
        plugin_name = self.tr("Atlas Grid Generator", "Atlas-Gittergenerator")

        self.action = QAction(QIcon(icon_path), plugin_name, self.iface.mainWindow())
        self.action.setToolTip(
            self.tr(
                "Regular grid generator for atlas workflows",
                "Regelmäßiger Gittergenerator für Atlas-Workflows"
            )
        )
        self.action.triggered.connect(self.show_dialog)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu(plugin_name, self.action)

    def unload(self):
        plugin_name = self.tr("Atlas Grid Generator", "Atlas-Gittergenerator")
        self.iface.removeToolBarIcon(self.action)
        self.iface.removePluginMenu(plugin_name, self.action)

    def show_dialog(self):
        dialog = QDialog(self.iface.mainWindow())
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
                "NO_LAYER"
            )
            self.layer_combo.setEnabled(False)

        layout.addWidget(self.layer_combo)

        self.selected_only_checkbox = QCheckBox(
            self.tr(
                "Create grid only for selected features",
                "Nur für ausgewählte Objekte ein Gitter erstellen"
            )
        )
        self.selected_only_checkbox.setToolTip(
            self.tr(
                "Use only selected features from the chosen layer.",
                "Verwendet nur ausgewählte Objekte des gewählten Layers."
            )
        )
        layout.addWidget(self.selected_only_checkbox)

        layout.addWidget(QLabel(self.tr("Select scale:", "Maßstab wählen:")))
        self.scale_combo = QComboBox()
        self.scale_options = [
            "1:500", "1:750", "1:1000", "1:1500", "1:2000", "1:2500", "1:3000",
            "1:4000", "1:5000", "1:7500", "1:10000", "1:15000", "1:20000", "1:25000"
        ]
        for scale in self.scale_options:
            self.scale_combo.addItem(scale, scale)
        layout.addWidget(self.scale_combo)

        scale_input_layout = QHBoxLayout()
        self.custom_scale_checkbox = QCheckBox(
            self.tr("Or enter scale manually:", "Oder Maßstab eingeben:")
        )
        self.scale_input = QLineEdit()
        self.scale_input.setPlaceholderText(self.tr("e.g. 3000", "z. B. 3000"))
        self.scale_input.setEnabled(False)

        scale_input_layout.addWidget(self.custom_scale_checkbox)
        scale_input_layout.addWidget(QLabel("1:"))
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
        self.manual_size_checkbox.stateChanged.connect(self.toggle_manual_size_mode)

        run_button = QPushButton(self.tr("Create grid", "Gitter erstellen"))
        run_button.clicked.connect(lambda: self.generate_grid(dialog))
        layout.addWidget(run_button)

        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        dialog.setLayout(layout)
        dialog.setMinimumWidth(390)
        dialog.exec()

    def toggle_scale_mode(self):
        is_custom = self.custom_scale_checkbox.isChecked()
        self.scale_combo.setEnabled(not is_custom)
        self.scale_input.setEnabled(is_custom)

    def toggle_manual_size_mode(self):
        is_manual = self.manual_size_checkbox.isChecked()
        self.manual_width.setEnabled(is_manual)
        self.manual_height.setEnabled(is_manual)
        self.paper_combo.setEnabled(not is_manual)

    def get_processing_crs(self, layer):
        source_crs = layer.crs()
        extent = layer.extent()

        if not source_crs.isValid() or extent.isEmpty():
            return QgsCoordinateReferenceSystem("EPSG:25832")

        if source_crs.isGeographic():
            center_x = (extent.xMinimum() + extent.xMaximum()) / 2.0
            center_y = (extent.yMinimum() + extent.yMaximum()) / 2.0
            zone = int((center_x + 180) / 6) + 1
            epsg_code = 32600 + zone if center_y >= 0 else 32700 + zone
            return QgsCoordinateReferenceSystem(f"EPSG:{epsg_code}")

        return source_crs

    def get_column_label(self, index):
        result = ""
        index -= 1
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    def get_grid_dimensions_mm(self, orientation, paper_size):
        height_mm, width_mm = self.paper_sizes_mm[paper_size]
        if orientation == "landscape":
            return width_mm, height_mm
        return height_mm, width_mm

    def get_scale_value(self, parent):
        if self.custom_scale_checkbox.isChecked():
            try:
                user_scale = int(self.scale_input.text().strip())
                if user_scale <= 0:
                    raise ValueError
                return user_scale
            except Exception:
                QMessageBox.warning(
                    parent,
                    self.tr("Error", "Fehler"),
                    self.tr("Please enter a valid scale value.", "Bitte eine gültige Maßstabzahl eingeben.")
                )
                return None

        scale_label = self.scale_combo.currentData().replace("1:", "")
        return int(scale_label)

    def get_grid_size_mm(self, parent, orientation, paper_size):
        if self.manual_size_checkbox.isChecked():
            try:
                grid_width_mm = float(self.manual_width.text().replace(",", "."))
                grid_height_mm = float(self.manual_height.text().replace(",", "."))
                if grid_width_mm <= 0 or grid_height_mm <= 0:
                    raise ValueError
                size_string = f"{int(grid_width_mm)}x{int(grid_height_mm)}mm"
                return grid_width_mm, grid_height_mm, size_string
            except Exception:
                QMessageBox.warning(
                    parent,
                    self.tr("Error", "Fehler"),
                    self.tr(
                        "Please enter valid width and height values.",
                        "Bitte gültige Werte für Breite und Höhe eingeben."
                    )
                )
                return None, None, None

        grid_width_mm, grid_height_mm = self.get_grid_dimensions_mm(orientation, paper_size)
        return grid_width_mm, grid_height_mm, paper_size

    def build_output_layer_name(self, layer_name, scale, orientation, size_string, selected_only):
        layer_base_name = layer_name.replace(" ", "_").replace(":", "_").replace("/", "_")
        source_mode = "selected" if selected_only else "layer"
        base_name = f"Gitter_1:{scale}_{orientation}_{size_string}_{source_mode}_{layer_base_name}"

        existing_names = [l.name() for l in QgsProject.instance().mapLayers().values()]
        counter = 1
        final_name = f"{base_name}_{counter:02d}"

        while final_name in existing_names:
            counter += 1
            final_name = f"{base_name}_{counter:02d}"

        return final_name

    def rect_to_source_polygon(self, rect, to_source_transform=None):
        points = [
            QgsPointXY(rect.xMinimum(), rect.yMinimum()),
            QgsPointXY(rect.xMaximum(), rect.yMinimum()),
            QgsPointXY(rect.xMaximum(), rect.yMaximum()),
            QgsPointXY(rect.xMinimum(), rect.yMaximum())
        ]

        if to_source_transform is not None:
            points = [to_source_transform.transform(pt) for pt in points]

        points.append(points[0])
        return QgsGeometry.fromPolygonXY([points])

    def add_grid_features(self, raw_cells, grid_layer, provider, to_source_transform=None):
        if not raw_cells:
            return 0

        sorted_cells = sorted(raw_cells, key=lambda item: (-item[7], item[6]))
        max_bottom_row = max(item[4] for item in raw_cells)

        new_features = []
        for serial, cell in enumerate(sorted_cells, start=1):
            x_min, y_min, x_max, y_max, row_from_bottom, col, _, _ = cell
            rect = QgsRectangle(x_min, y_min, x_max, y_max)

            feat = QgsFeature(grid_layer.fields())
            feat.setGeometry(self.rect_to_source_polygon(rect, to_source_transform))

            row_from_top = max_bottom_row - row_from_bottom + 1
            feat.setAttribute("grid", f"{self.get_column_label(col)}{row_from_top}")
            feat.setAttribute("serial", serial)

            new_features.append(feat)

        provider.addFeatures(new_features)
        return len(new_features)

    def generate_grid(self, dialog):
        layer_name = self.layer_combo.currentData()
        orientation = self.format_combo.currentData()
        paper_size = self.paper_combo.currentData()
        selected_only = self.selected_only_checkbox.isChecked()

        if layer_name == "NO_LAYER":
            QMessageBox.warning(
                dialog,
                self.tr("Error", "Fehler"),
                self.tr("No valid vector layer was found.", "Es wurde kein gültiger Vektor-Layer gefunden.")
            )
            return

        scale = self.get_scale_value(dialog)
        if scale is None:
            return

        grid_width_mm, grid_height_mm, size_string = self.get_grid_size_mm(dialog, orientation, paper_size)
        if grid_width_mm is None or grid_height_mm is None:
            return

        grid_width = (grid_width_mm / 1000.0) * scale
        grid_height = (grid_height_mm / 1000.0) * scale

        layers = QgsProject.instance().mapLayersByName(layer_name)
        if not layers:
            QMessageBox.warning(
                dialog,
                self.tr("Error", "Fehler"),
                self.tr("The selected layer was not found.", "Der ausgewählte Layer wurde nicht gefunden.")
            )
            return

        layer = layers[0]
        if not isinstance(layer, QgsVectorLayer) or not layer.isValid():
            QMessageBox.warning(
                dialog,
                self.tr("Error", "Fehler"),
                self.tr("The selected layer is invalid.", "Der ausgewählte Layer ist ungültig.")
            )
            return

        if selected_only:
            if layer.selectedFeatureCount() == 0:
                QMessageBox.warning(
                    dialog,
                    self.tr("Error", "Fehler"),
                    self.tr(
                        "Selected-features mode is enabled, but no features are selected.",
                        "Die Option für ausgewählte Objekte ist aktiviert, aber im Layer ist nichts selektiert."
                    )
                )
                return
            source_features = list(layer.getSelectedFeatures())
        else:
            source_features = list(layer.getFeatures())

        if not source_features:
            QMessageBox.information(
                dialog,
                self.tr("Information", "Hinweis"),
                self.tr("No features were found for processing.", "Keine Features zum Verarbeiten gefunden.")
            )
            return

        source_crs = layer.crs()
        processing_crs = self.get_processing_crs(layer)
        transform_context = QgsProject.instance().transformContext()

        needs_transform = source_crs.authid() != processing_crs.authid()
        to_processing = QgsCoordinateTransform(source_crs, processing_crs, transform_context) if needs_transform else None
        to_source = QgsCoordinateTransform(processing_crs, source_crs, transform_context) if needs_transform else None

        progress = QProgressDialog(
            self.tr("Preparing geometries...", "Geometrien werden vorbereitet..."),
            self.tr("Cancel", "Abbrechen"),
            0,
            100,
            self.iface.mainWindow()
        )
        progress.setWindowTitle(self.tr("Please wait", "Bitte warten"))
        progress.setModal(True)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        progress.show()
        QApplication.processEvents()

        transformed_geometries = []
        total = len(source_features)

        try:
            for i, feature in enumerate(source_features):
                if progress.wasCanceled():
                    progress.close()
                    return

                geom = QgsGeometry(feature.geometry())
                if geom.isEmpty():
                    continue

                if needs_transform:
                    geom.transform(to_processing)

                if not geom.isEmpty():
                    transformed_geometries.append(geom)

                prep_percent = int(((i + 1) / total) * 25)
                progress.setValue(prep_percent)
                QApplication.processEvents()

        except Exception as e:
            progress.close()
            QMessageBox.critical(
                dialog,
                self.tr("Error", "Fehler"),
                self.tr(
                    f"Error during geometry transformation:\n{str(e)}",
                    f"Fehler bei der Geometrietransformation:\n{str(e)}"
                )
            )
            return

        if not transformed_geometries:
            progress.close()
            QMessageBox.information(
                dialog,
                self.tr("Information", "Hinweis"),
                self.tr("No valid geometries were found in the layer.", "Keine gültigen Geometrien im Layer.")
            )
            return

        first_bbox = transformed_geometries[0].boundingBox()
        xmin = first_bbox.xMinimum()
        xmax = first_bbox.xMaximum()
        ymin = first_bbox.yMinimum()
        ymax = first_bbox.yMaximum()

        for geom in transformed_geometries[1:]:
            bbox = geom.boundingBox()
            xmin = min(xmin, bbox.xMinimum())
            xmax = max(xmax, bbox.xMaximum())
            ymin = min(ymin, bbox.yMinimum())
            ymax = max(ymax, bbox.yMaximum())

        offset = 10.0
        xmin -= offset
        xmax += offset
        ymin -= offset
        ymax += offset

        grid_layer_name = self.build_output_layer_name(
            layer.name(), scale, orientation, size_string, selected_only
        )

        grid_layer = QgsVectorLayer(f"Polygon?crs={source_crs.authid()}", grid_layer_name, "memory")
        provider = grid_layer.dataProvider()
        provider.addAttributes([
            QgsField("grid", QMetaType.Type.QString),
            QgsField("serial", QMetaType.Type.Int)
        ])
        grid_layer.updateFields()

        self.worker = GridGeneratorThread(
            transformed_geometries=transformed_geometries,
            grid_width=grid_width,
            grid_height=grid_height,
            xmin=xmin,
            xmax=xmax,
            ymin=ymin,
            ymax=ymax
        )

        def on_progress(val):
            mapped_val = 25 + int(val * 0.75)
            progress.setLabelText(self.tr("Creating grid cells...", "Gitterzellen werden erzeugt..."))
            progress.setValue(mapped_val)
            QApplication.processEvents()

        def on_failed(message):
            progress.close()
            QMessageBox.critical(
                dialog,
                self.tr("Error", "Fehler"),
                self.tr(
                    f"Grid generation failed:\n{message}",
                    f"Gittererzeugung fehlgeschlagen:\n{message}"
                )
            )

        def on_cancelled():
            progress.close()
            QMessageBox.information(
                dialog,
                self.tr("Cancelled", "Abgebrochen"),
                self.tr("Grid generation was cancelled.", "Die Gittererzeugung wurde abgebrochen.")
            )

        def on_finished(raw_cells):
            progress.close()

            if not raw_cells:
                QMessageBox.information(
                    dialog,
                    self.tr("Information", "Hinweis"),
                    self.tr(
                        "No matching grid cells were found.",
                        "Es wurden keine passenden Gitterzellen gefunden."
                    )
                )
                return

            try:
                count = self.add_grid_features(
                    raw_cells=raw_cells,
                    grid_layer=grid_layer,
                    provider=provider,
                    to_source_transform=to_source
                )

                grid_layer.updateExtents()

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
                label_settings.placement = Qgis.LabelPlacement.AroundPoint
                label_settings.enabled = True

                grid_layer.setLabeling(QgsVectorLayerSimpleLabeling(label_settings))
                grid_layer.setLabelsEnabled(True)

                QgsProject.instance().addMapLayer(grid_layer)
                grid_layer.triggerRepaint()
                self.iface.layerTreeView().refreshLayerSymbology(grid_layer.id())

                if selected_only:
                    msg = self.tr(
                        f"{count} grid cells were created only for the selected features.",
                        f"{count} Gitterzellen wurden nur für die ausgewählten Objekte erstellt."
                    )
                else:
                    msg = self.tr(
                        f"{count} grid cells were created for the entire layer.",
                        f"{count} Gitterzellen wurden für den gesamten Layer erstellt."
                    )

                QMessageBox.information(
                    dialog,
                    self.tr("Done", "Fertig"),
                    msg
                )
                dialog.close()

            except Exception as e:
                QMessageBox.critical(
                    dialog,
                    self.tr("Error", "Fehler"),
                    self.tr(
                        f"Error while writing the result layer:\n{str(e)}",
                        f"Fehler beim Schreiben des Ergebnislayers:\n{str(e)}"
                    )
                )

        progress.canceled.connect(lambda: self.worker.cancel())
        self.worker.progressChanged.connect(on_progress)
        self.worker.failed.connect(on_failed)
        self.worker.cancelled.connect(on_cancelled)
        self.worker.finished.connect(on_finished)
        self.worker.start()