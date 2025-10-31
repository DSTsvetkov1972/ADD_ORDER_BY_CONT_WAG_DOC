
cell_khaki = 'k'
cell_white = 'h'

header = {
    (1, 3): ('№', cell_khaki),
    (1, 4): ('№ Вагона', cell_khaki),
    (1, 5): ('Номер накладной', cell_khaki),
    (1, 6): ('№ Контейнера', cell_khaki),
    (1, 7): ('Ст. отправления', cell_khaki),
    (1, 8): ('Ст. назначения', cell_khaki),
    (1, 12): ('Дата отправки', cell_khaki),       
    (1, 13): ('использование пути', cell_khaki),
    (1, 14): (' тариф груженый', cell_khaki)
    }

for k, (column_name, column_fill) in header.items():
    print(k, column_name, column_fill)