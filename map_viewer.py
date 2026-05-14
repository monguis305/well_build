import sys
import os
import numpy as np
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout,
                             QWidget, QPushButton, QFileDialog, QLabel,
                             QHBoxLayout, QStatusBar, QSizePolicy)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from scipy.interpolate import RegularGridInterpolator
import xtgeo


class MapViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IRAP ASCII Map Viewer")

        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        self.setGeometry(0, 0, screen_geometry.width(), screen_geometry.height())

        self.data = None
        self.filepath = None
        self.current_data = None
        self.real_extent = None
        self.display_extent = None
        self.nrows = 0
        self.ncols = 0
        self.cbar = None
        self.im = None

        # Данные скважин
        self.wells = None
        self.well_points = None

        # Режим перемещения скважин
        self.move_wells_mode = False
        self.wells_dragging = False
        self.wells_drag_start = None
        self.wells_original_coords = None

        # Интерполятор для снятия значений с карты
        self.interpolator = None

        # Диаграмма справа
        self.chart_figure = None
        self.chart_canvas = None
        self.chart_ax = None

        self.initUI()

    def initUI(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: white;
            }
            QWidget {
                background-color: white;
            }
        """)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # ===== Левая часть — карта =====
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(3)

        # Кнопки
        btn_layout = QHBoxLayout()
        self.btn_open = QPushButton("Открыть карту")
        self.btn_open.clicked.connect(self.open_file)
        btn_layout.addWidget(self.btn_open)

        self.btn_load_wells = QPushButton("Загрузить скважины")
        self.btn_load_wells.clicked.connect(self.load_wells)
        btn_layout.addWidget(self.btn_load_wells)

        self.btn_clear_wells = QPushButton("Убрать скважины")
        self.btn_clear_wells.clicked.connect(self.clear_wells)
        btn_layout.addWidget(self.btn_clear_wells)

        self.btn_clear = QPushButton("Очистить")
        self.btn_clear.clicked.connect(self.clear_plot)
        btn_layout.addWidget(self.btn_clear)

        btn_layout.addStretch()
        left_layout.addLayout(btn_layout)

        # Трансформации
        transform_layout = QHBoxLayout()
        transform_layout.addWidget(QLabel("Трансформации:"))

        self.btn_flip_v = QPushButton("↕ Отр. верт.")
        self.btn_flip_v.clicked.connect(self.flip_vertical)
        transform_layout.addWidget(self.btn_flip_v)

        self.btn_flip_h = QPushButton("↔ Отр. гор.")
        self.btn_flip_h.clicked.connect(self.flip_horizontal)
        transform_layout.addWidget(self.btn_flip_h)

        self.btn_rot_left = QPushButton("↺ -90°")
        self.btn_rot_left.clicked.connect(lambda: self.rotate(-90))
        transform_layout.addWidget(self.btn_rot_left)

        self.btn_rot_right = QPushButton("↻ +90°")
        self.btn_rot_right.clicked.connect(lambda: self.rotate(90))
        transform_layout.addWidget(self.btn_rot_right)

        self.btn_rot_180 = QPushButton("180°")
        self.btn_rot_180.clicked.connect(lambda: self.rotate(180))
        transform_layout.addWidget(self.btn_rot_180)

        self.btn_reset = QPushButton("Сброс")
        self.btn_reset.clicked.connect(self.reset_view)
        transform_layout.addWidget(self.btn_reset)

        # Кнопка перемещения скважин
        self.btn_move_wells = QPushButton("⇱ Двигать скважины")
        self.btn_move_wells.setCheckable(True)
        self.btn_move_wells.toggled.connect(self.toggle_move_wells)
        transform_layout.addWidget(self.btn_move_wells)

        transform_layout.addStretch()
        left_layout.addLayout(transform_layout)

        # Область карты — КВАДРАТНАЯ
        screen = QApplication.primaryScreen()
        screen_height = screen.availableGeometry().height()
        square_size = int(screen_height * 0.88)

        self.figure = Figure(figsize=(square_size / 100, square_size / 100), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setFixedSize(square_size, square_size)

        self.canvas.mpl_connect('scroll_event', self.on_scroll)
        self.canvas.mpl_connect('button_press_event', self.on_press)
        self.canvas.mpl_connect('button_release_event', self.on_release)
        self.canvas.mpl_connect('motion_notify_event', self.on_motion)

        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Нажмите 'Открыть карту'")
        self.ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center',
                     fontsize=14, color='white')
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        left_layout.addWidget(self.canvas)

        # Инфо-лейбл
        self.info_label = QLabel("")
        self.info_label.setMaximumHeight(20)
        self.info_label.setStyleSheet("background-color: #e8e8e8; padding: 2px; font-size: 11px;")
        left_layout.addWidget(self.info_label)

        main_layout.addWidget(left_widget)

        # ===== Правая часть — диаграмма =====
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        self.chart_figure = Figure(figsize=(5, 8), dpi=100)
        self.chart_canvas = FigureCanvas(self.chart_figure)
        self.chart_canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.chart_ax = self.chart_figure.add_subplot(111)
        self.chart_ax.set_title("Значения в скважинах", fontsize=11)
        self.chart_ax.set_xlabel("Скважина", fontsize=9)
        self.chart_ax.set_ylabel("Значение", fontsize=9)
        self.chart_ax.tick_params(labelsize=8)
        right_layout.addWidget(self.chart_canvas)

        main_layout.addWidget(right_widget, 1)  # stretch factor 1 — занимает оставшееся место

        # Статусная строка
        self.status_bar = QStatusBar()
        self.status_bar.setMaximumHeight(20)
        self.status_bar.setStyleSheet("font-size: 11px;")
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов к работе | Колёсико — зум | Зажатая мышь — панорама")

    # ===== Методы для перемещения скважин =====

    def toggle_move_wells(self, checked):
        """Включение/выключение режима перемещения скважин"""
        self.move_wells_mode = checked
        if checked:
            self.status_bar.showMessage("Режим перемещения скважин ВКЛЮЧЕН | Зажмите мышь и двигайте")
        else:
            self.status_bar.showMessage("Режим перемещения скважин выключен")

    def on_press(self, event):
        if event.inaxes != self.ax:
            return

        # Если режим перемещения скважин включён
        if self.move_wells_mode and self.wells is not None and event.button == 1:
            self.wells_dragging = True
            self.wells_drag_start = (event.xdata, event.ydata)
            self.wells_original_coords = self.wells[['x', 'y']].copy()
            return

        # Обычное панорамирование
        if event.button == 1:
            self._dragging = True
            self._pan_start = (event.xdata, event.ydata)
            self._xlim_start = self.ax.get_xlim()
            self._ylim_start = self.ax.get_ylim()

    def on_release(self, event):
        if self.wells_dragging:
            self.wells_dragging = False
            # Обновляем значения data после перемещения
            self.update_wells_data()
            self.update_chart()
            self.draw_map()

        self._dragging = False

    def on_motion(self, event):
        # Перемещение скважин
        if self.wells_dragging and event.inaxes == self.ax:
            if event.xdata is None or event.ydata is None:
                return

            dx = event.xdata - self.wells_drag_start[0]
            dy = event.ydata - self.wells_drag_start[1]

            self.wells['x'] = self.wells_original_coords['x'] + dx
            self.wells['y'] = self.wells_original_coords['y'] + dy

            # Обновляем только скважины на карте (быстро)
            self.draw_map()
            self.update_chart()
            return

        # Обычное панорамирование
        if not self._dragging or event.inaxes != self.ax:
            return

        if event.xdata is None or event.ydata is None:
            return

        dx = self._pan_start[0] - event.xdata
        dy = self._pan_start[1] - event.ydata

        self.ax.set_xlim(self._xlim_start[0] + dx, self._xlim_start[1] + dx)
        self.ax.set_ylim(self._ylim_start[0] + dy, self._ylim_start[1] + dy)
        self.canvas.draw_idle()

    # ===== Методы для работы со значениями скважин =====

    def update_wells_data(self):
        """Обновление столбца data в wells — снятие значений с карты"""
        if self.wells is None or self.current_data is None:
            return

        # Создаём интерполятор, если ещё нет
        if self.interpolator is None:
            self.create_interpolator()

        if self.interpolator is None:
            self.status_bar.showMessage("Не удалось создать интерполятор")
            return

        # Если столбца data нет — создаём
        if 'data' not in self.wells.columns:
            self.wells['data'] = np.nan

        # Снимаем значения с карты для ВСЕХ скважин (или только для тех, где data пустая)
        for idx in self.wells.index:
            # Снимаем значение, если data пустая или если скважины двигались
            x = self.wells.at[idx, 'x']
            y = self.wells.at[idx, 'y']
            try:
                val = self.interpolator([y, x])[0]  # ВАЖНО: порядок (y, x)!
                self.wells.at[idx, 'data'] = val
            except Exception as e:
                print(f"Ошибка интерполяции для {self.wells.at[idx, 'name']}: {e}")
                pass

    def create_interpolator(self):
        """Создание интерполятора по данным карты"""
        if self.current_data is None:
            return

        xmin, xmax = self.real_extent[0], self.real_extent[1]
        ymin, ymax = self.real_extent[2], self.real_extent[3]

        # Создаём сетки координат
        x_coords = np.linspace(xmin, xmax, self.current_data.shape[1])
        y_coords = np.linspace(ymin, ymax, self.current_data.shape[0])

        # Заполняем NaN ближайшими значениями для интерполяции
        data_filled = self.current_data.copy()
        mask = np.isnan(data_filled)
        if mask.any():
            from scipy.interpolate import griddata
            yy, xx = np.meshgrid(y_coords, x_coords, indexing='ij')
            points = np.column_stack((xx[~mask], yy[~mask]))
            values = data_filled[~mask]
            data_filled[mask] = griddata(points, values, (xx[mask], yy[mask]), method='nearest')

        self.interpolator = RegularGridInterpolator(
            (y_coords, x_coords),
            data_filled,
            bounds_error=False,
            fill_value=np.nan
        )

    def update_chart(self):
        """Обновление столбчатой диаграммы (X = prod, Y = data, подписи внутри и сверху)"""
        if self.chart_ax is None:
            return

        self.chart_ax.clear()

        if self.wells is None or len(self.wells) == 0:
            self.chart_ax.set_title("Нет скважин", fontsize=11)
            self.chart_canvas.draw_idle()
            return

        # Обновляем значения
        self.update_wells_data()

        # Сортируем по prod для наглядности
        wells_sorted = self.wells.sort_values('prod')

        prod_values = wells_sorted['prod'].values
        data_values = wells_sorted['data'].values
        names = wells_sorted['name'].values

        # Если prod пустой — используем индексы
        if np.all(np.isnan(prod_values)):
            x_labels = names
            x_positions = range(len(names))
        else:
            x_labels = [f'{p:.1f}' for p in prod_values]
            x_positions = range(len(names))

        # Цвета столбцов
        colors = ['steelblue' if v is not None and not np.isnan(v) else 'gray' for v in data_values]

        bars = self.chart_ax.bar(x_positions, data_values, color=colors,
                                 edgecolor='black', linewidth=0.5, width=0.7)

        # Настройка оси X
        self.chart_ax.set_xticks(x_positions)
        self.chart_ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=7)
        self.chart_ax.set_xlabel("Дебит (prod)", fontsize=9)
        self.chart_ax.set_ylabel("Значение с карты", fontsize=9)
        self.chart_ax.set_title("Значения в скважинах", fontsize=11)
        self.chart_ax.tick_params(labelsize=8)
        self.chart_ax.grid(axis='y', alpha=0.3)

        # Подписи
        for bar, name, val in zip(bars, names, data_values):
            if not np.isnan(val) and bar.get_height() > 0:
                # Название скважины внутри столбца (вертикально)
                self.chart_ax.text(bar.get_x() + bar.get_width() / 2,
                                   bar.get_height() / 2,
                                   name,
                                   ha='center', va='center',
                                   fontsize=6, color='white', fontweight='bold',
                                   rotation=90)

                # Значение НАД столбцом
                self.chart_ax.text(bar.get_x() + bar.get_width() / 2,
                                   bar.get_height() + (max(data_values) * 0.01),  # чуть выше столбца
                                   f'{val:.2f}',
                                   ha='center', va='bottom',
                                   fontsize=7, color='black', fontweight='bold')

        self.chart_figure.tight_layout()
        self.chart_canvas.draw_idle()
    # ===== Загрузка скважин (обновлённая) =====

    def load_wells(self):
        """Загрузка координат скважин из Excel (A: имя, B: X, C: Y, D: data, E: prod)"""
        if self.data is None:
            self.status_bar.showMessage("Сначала загрузите карту!")
            return

        filepath, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл со скважинами", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )

        if not filepath:
            return

        try:
            df = pd.read_excel(filepath, header=None)

            # Если есть 5 столбцов — используем, иначе заполняем пустыми
            if df.shape[1] >= 5:
                df = df.iloc[:, [0, 1, 2, 3, 4]]
            elif df.shape[1] == 4:
                df = df.iloc[:, [0, 1, 2, 3]]
                df[4] = np.nan
            else:
                df = df.iloc[:, [0, 1, 2]]
                df[3] = np.nan
                df[4] = np.nan

            df.columns = ['name', 'x', 'y', 'data', 'prod']

            df['name'] = df['name'].astype(str)
            df['x'] = pd.to_numeric(df['x'], errors='coerce')
            df['y'] = pd.to_numeric(df['y'], errors='coerce')
            df['data'] = pd.to_numeric(df['data'], errors='coerce')
            df['prod'] = pd.to_numeric(df['prod'], errors='coerce')

            df = df.dropna(subset=['x', 'y'])

            self.wells = df

            # Снимаем значения с карты
            self.interpolator = None
            self.update_wells_data()

            if self.current_data is not None:
                self.draw_map()
            self.update_chart()

            self.status_bar.showMessage(f"Загружено скважин: {len(df)}")

        except Exception as e:
            import traceback
            self.status_bar.showMessage(f"Ошибка загрузки скважин: {str(e)}")
            traceback.print_exc()

    # ===== Остальные методы (без изменений) =====

    def clear_wells(self):
        self.wells = None
        self.well_points = None
        self.interpolator = None
        if self.current_data is not None:
            self.draw_map()
        self.update_chart()
        self.status_bar.showMessage("Скважины убраны")

    def on_scroll(self, event):
        if event.inaxes != self.ax:
            return

        scale_factor = 0.8 if event.button == 'up' else 1.25
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        xdata, ydata = event.xdata, event.ydata

        if xdata is None or ydata is None:
            return

        new_xlim = [xdata - (xdata - xlim[0]) * scale_factor,
                    xdata + (xlim[1] - xdata) * scale_factor]
        new_ylim = [ydata - (ydata - ylim[0]) * scale_factor,
                    ydata + (ylim[1] - ydata) * scale_factor]

        self.ax.set_xlim(new_xlim)
        self.ax.set_ylim(new_ylim)
        self.canvas.draw_idle()

    def load_and_display(self, filepath):
        try:
            surf = xtgeo.surface_from_file(filepath, fformat='irap_ascii')
            self.data = surf.values.copy()

            nodata = surf.undef if hasattr(surf, 'undef') else 9999900.0
            self.data = np.where(np.isclose(self.data, nodata), np.nan, self.data)

            xmin, xmax = surf.xmin, surf.xmax
            ymin, ymax = surf.ymin, surf.ymax
            self.real_extent = [xmin, xmax, ymin, ymax]

            dx_range = xmax - xmin
            dy_range = ymax - ymin
            x_center = (xmin + xmax) / 2
            y_center = (ymin + ymax) / 2

            self.display_extent = [
                x_center - dx_range * 2.5,
                x_center + dx_range * 2.5,
                y_center - dy_range * 2.5,
                y_center + dy_range * 2.5,
            ]

            self.nrows, self.ncols = surf.nrow, surf.ncol
            self.current_data = self.data.copy()

            # Сброс интерполятора при новой карте
            self.interpolator = None

            self.draw_map()
            self.status_bar.showMessage(f"Загружен: {filepath}")

        except Exception as e:
            import traceback
            self.status_bar.showMessage(f"Ошибка: {str(e)}")
            traceback.print_exc()

    def draw_map(self):
        if self.current_data is None:
            return

        self.figure.clear()
        self.ax = self.figure.add_axes([0.05, 0.05, 0.87, 0.9])
        self.ax.tick_params(axis='both', labelsize=5)

        if getattr(self, '_rotated_90', False):
            data_xmin, data_xmax = self.real_extent[2], self.real_extent[3]
            data_ymin, data_ymax = self.real_extent[0], self.real_extent[1]

            self.im = self.ax.imshow(self.current_data, cmap='viridis',
                                     extent=[data_xmin, data_xmax, data_ymin, data_ymax],
                                     aspect='auto', origin='lower')
            self.ax.set_xlabel("Y координата", fontsize=6)
            self.ax.set_ylabel("X координата", fontsize=6)
        else:
            data_xmin, data_xmax = self.real_extent[0], self.real_extent[1]
            data_ymin, data_ymax = self.real_extent[2], self.real_extent[3]

            self.im = self.ax.imshow(self.current_data, cmap='viridis',
                                     extent=self.real_extent,
                                     aspect='auto', origin='lower')
            self.ax.set_xlabel("X координата", fontsize=6)
            self.ax.set_ylabel("Y координата", fontsize=6)

        # Рисуем скважины
        if self.wells is not None and len(self.wells) > 0:
            if getattr(self, '_rotated_90', False):
                self.ax.scatter(self.wells['y'], self.wells['x'],
                              c='red', marker='o', s=50, edgecolors='black',
                              linewidth=1, zorder=5, label='Скважины')
                for _, row in self.wells.iterrows():
                    self.ax.annotate(str(row['name']),
                                   xy=(row['y'], row['x']),
                                   xytext=(5, 5), textcoords='offset points',
                                   fontsize=8, color='black', fontweight='bold',
                                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.7))
            else:
                self.ax.scatter(self.wells['x'], self.wells['y'],
                              c='red', marker='o', s=50, edgecolors='black',
                              linewidth=1, zorder=5, label='Скважины')
                for _, row in self.wells.iterrows():
                    self.ax.annotate(str(row['name']),
                                   xy=(row['x'], row['y']),
                                   xytext=(5, 5), textcoords='offset points',
                                   fontsize=8, color='black', fontweight='bold',
                                   bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.7))

        self.cbar = self.figure.colorbar(self.im, ax=self.ax, fraction=0.046, pad=0.02)
        self.ax.set_title(f"Карта: {os.path.basename(self.filepath)}", fontsize=10)
        self.ax.grid(True, alpha=0.3, linestyle='--')
        self.ax.set_xlim(self.ax.get_xlim())
        self.ax.set_ylim(self.ax.get_ylim())

        if self.wells is not None and len(self.wells) > 0:
            self.ax.legend(loc='upper right', fontsize=8)

        self.canvas.draw_idle()

        valid = self.current_data[~np.isnan(self.current_data)]
        if len(valid) > 0:
            self.info_label.setText(
                f"Размер: {self.nrows}×{self.ncols} | "
                f"Min: {np.nanmin(valid):.3f} | Max: {np.nanmax(valid):.3f} | "
                f"Скважин: {len(self.wells) if self.wells is not None else 0}"
            )

    def update_map_data(self):
        if self.im is None:
            self.draw_map()
            return

        self.im.set_data(self.current_data)
        vmin = np.nanmin(self.current_data)
        vmax = np.nanmax(self.current_data)
        if np.isnan(vmin) or np.isnan(vmax):
            vmin, vmax = 0, 1
        self.im.set_clim(vmin, vmax)
        self.canvas.draw_idle()

    def flip_vertical(self):
        if self.current_data is not None:
            self.current_data = np.flipud(self.current_data)
            self.interpolator = None
            self.update_map_data()
            self.update_chart()
            self.status_bar.showMessage("Отражено по вертикали")

    def flip_horizontal(self):
        if self.current_data is not None:
            self.current_data = np.fliplr(self.current_data)
            self.interpolator = None
            self.update_map_data()
            self.update_chart()
            self.status_bar.showMessage("Отражено по горизонтали")

    def rotate(self, degrees):
        if self.current_data is not None:
            k = (degrees // 90) % 4
            self.current_data = np.rot90(self.current_data, k=k)
            self._rotated_90 = (k % 2 == 1)
            self.interpolator = None
            self.draw_map()
            self.update_chart()
            self.status_bar.showMessage(f"Повёрнуто на {degrees}°")

    def reset_view(self):
        if self.data is not None:
            self.current_data = self.data.copy()
            self._rotated_90 = False
            self.interpolator = None
            self.draw_map()
            self.update_chart()
            self.status_bar.showMessage("Вид сброшен")

    def open_file(self):
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл карты", "",
            "All Files (*);;Text Files (*.txt *.asc *.dat *.gri)"
        )
        if filepath:
            self.filepath = filepath
            self.load_and_display(filepath)

    def clear_plot(self):
        self.data = None
        self.current_data = None
        self.im = None
        self.cbar = None
        self.filepath = None
        self.wells = None
        self.interpolator = None
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Нажмите 'Открыть карту'")
        self.ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center',
                     fontsize=14, color='gray')
        self.ax.set_xticks([])
        self.ax.set_yticks([])
        self.canvas.draw_idle()
        self.update_chart()
        self.info_label.setText("")
        self.status_bar.showMessage("Готов к работе | Колёсико — зум | Зажатая мышь — панорама")


def main():
    app = QApplication(sys.argv)
    viewer = MapViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
