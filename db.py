from PyQt6.QtSql import QSqlDatabase

DB_PATH = "tech.db"

def create_connection():
    db = QSqlDatabase.addDatabase("QSQLITE")
    db.setDatabaseName(DB_PATH)
    if not db.open():
        raise Exception("Не удалось открыть базу данных")
    return db