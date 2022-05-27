import os
import time
import pandas as pd
import shutil
import platform
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class ExcelFileHandler:
    __excel_files_directory = None
    __excel_files = None
    __master_file = None
    __processed_path = None

    def __init__(self, excel_files_path, excel_master_file, processed_path):
        self.__excel_files_directory = excel_files_path
        self.__excel_files = list(Path(self.__excel_files_directory).glob("*.xlsx"))
        self.__master_file = excel_master_file
        self.__processed_path = processed_path

    def copy_to_master_file(self):
        df_total = pd.DataFrame()
        if len(self.__excel_files) == 0:
            return

        for excel_file in self.__excel_files:
            self.__move_excel(excel_file)

        self.__excel_files = list(Path(self.__processed_path).glob("*.xlsx"))

        for excel_file in self.__excel_files:
            file = pd.ExcelFile(excel_file)
            sheets = file.sheet_names
            for sheet in sheets:
                df = file.parse(sheet_name=sheet)
                df_total = df_total.append(df)

        df_total.to_excel(self.__master_file)

    def __move_excel(self, filename):
        shutil.move(filename, self.__processed_path)


# Object to manage archives that are not xlsx
class TrashHandler:
    __path_to_classify = None
    __invalid_path = None
    __trash_files = None

    def __init__(self, path_to_classify, trash_path):
        self.__path_to_classify = path_to_classify
        self.__invalid_path = trash_path
        self.__trash_files = list(Path(self.__path_to_classify).glob('*.*'))

    def remove_trash(self):
        for trash in self.__trash_files:
            shutil.move(trash, self.__invalid_path)


class PathHandler:
    __source_path = None
    __master_file_path = None
    __processed_Path = None
    __invalid_path = None

    def __init__(self, source):
        self.__source_path = source
        self.__processed_Path = self.__make_a_path("Processed")
        self.__invalid_path = self.__make_a_path("Not applicable")
        self.__master_file_path = self.__make_a_path("Master file")

    def __make_a_path(self, path_to_create):
        if not os.path.exists(self.__source_path + "\\" + path_to_create):
            new_dir = os.path.join(self.__source_path, path_to_create)
            os.mkdir(new_dir)
        return self.__source_path + "\\" + path_to_create

    def get_source(self):
        return self.__source_path

    def get_processed(self):
        return self.__processed_Path

    def get_invalid(self):
        return self.__invalid_path

    def get_master_file_path(self):
        return self.__master_file_path


def process_files():
    pass


if __name__ == '__main__':
    source_path = input("Enter a path: ")
    while not os.path.exists(source_path):
        print("Invalid path.")
        source_path = input("Enter a path: ")

    paths = PathHandler(source_path)
    if platform.system() == 'Windows':
        master_file = paths.get_master_file_path() + "\\Master_file.xlsx"
    else:
        master_file = paths.get_master_file_path() + "/Master_file.xlsx"

    excel_handler = ExcelFileHandler(paths.get_source(), master_file, paths.get_processed())
    excel_handler.copy_to_master_file()
    trash_handler = TrashHandler(paths.get_source(), paths.get_invalid())
    trash_handler.remove_trash()

    event_handler = FileSystemEventHandler()

    def on_created(event):
        time.sleep(3)
        inner_excel_handler = ExcelFileHandler(paths.get_source(), master_file, paths.get_processed())
        inner_excel_handler.copy_to_master_file()
        inner_trash_handler = TrashHandler(paths.get_source(), paths.get_invalid())
        inner_trash_handler.remove_trash()


    event_handler.on_created = on_created

    observer = Observer()
    observer.schedule(event_handler, paths.get_source(), recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
