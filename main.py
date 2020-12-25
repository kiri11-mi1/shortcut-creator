from tkinter import Tk
from tkinter.filedialog import askopenfile
from win32com.client import Dispatch

# Отключил отрисовку корневого графического окна
Tk().withdraw()

# Открытие диалогового окна для выбора файла
file = askopenfile(mode ='r', filetypes =[('Shortcuts', '*.url')])
rows = file.readlines()

params = {
    'path': file.name.replace('.url', '.lnk'),
    'WorkingDirectory': None,
    'URL': None,
    'IconFile': None,
}

file.close()

for row in rows:
    for key in params.keys():
        if key in row:
            start = row.find(key)
            # Прибаляем единицу, потому что следующий символ это '='
            end = start + len(key) + 1
            params[key] = row[end:-1]

path_to_exe = '\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe'


# С помощью метода Dispatch, обьявляем работу с Wscript (работа с ярлыками, реестром и прочей системной информацией в windows)
shell = Dispatch('WScript.Shell')

# Создаём ярлык.
shortcut = shell.CreateShortCut(params['path'])

# Путь к файлу, к которому делаем ярлык.
shortcut.Targetpath = params['WorkingDirectory'] + path_to_exe

# Путь к рабочей папке.
shortcut.WorkingDirectory = params['WorkingDirectory']

# Тырим иконку.
shortcut.IconLocation = params['IconFile']

# Добавляем аргумент для запуска игры
shortcut.Arguments = params['URL']

# Обязательное действо, сохраняем ярлык.
shortcut.save()
