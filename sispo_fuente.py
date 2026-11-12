"""
Fuente única del SISPO (sin Streamlit): el archivo se carga UNA vez en Administración y se guarda aquí, por compañía y periodo.
Mantenimiento y Operaciones lo leen de esta tabla; ninguno pide el archivo de nuevo.
"""
import io

TABLA = "sispo_archivo"


def ddl(motor: str) -> str:
    binario = "BYTEA" if motor == "postgres" else "BLOB"
    return f"CREATE TABLE IF NOT EXISTS {TABLA}(empresa TEXT, periodo INT, archivo TEXT, usuario TEXT, fecha TEXT, tamano INT, datos {binario})"


class Archivo(io.BytesIO):
    """El contenido guardado, con `name` y `getvalue()` como un archivo subido: lo leen igual `pd.read_excel` y las funciones que reciben `up`."""

    def __init__(self, contenido: bytes, nombre: str):
        super().__init__(contenido)
        self.name = nombre


def guardar(con, empresa: str, periodo: int, nombre: str, usuario: str, fecha: str, contenido: bytes):
    """Un archivo por compañía y periodo: cargar otro del mismo mes lo reemplaza."""
    con.execute(f"DELETE FROM {TABLA} WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    con.execute(f"INSERT INTO {TABLA}(empresa, periodo, archivo, usuario, fecha, tamano, datos) VALUES (?,?,?,?,?,?,?)",
                (empresa, int(periodo), nombre, usuario, fecha, len(contenido), contenido))
    con.commit()


def listar(leer):
    """Los archivos guardados, sin el contenido (que pesa)."""
    return leer(f"SELECT empresa, periodo, archivo, usuario, fecha, tamano FROM {TABLA} ORDER BY periodo DESC, empresa")


def abrir(leer, empresa: str, periodo: int):
    """El archivo guardado como `Archivo`, o None si no hay."""
    d = leer(f"SELECT archivo, datos FROM {TABLA} WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    if d.empty:
        return None
    return Archivo(bytes(d["datos"].iloc[0]), str(d["archivo"].iloc[0]))


def borrar(con, empresa: str, periodo: int):
    con.execute(f"DELETE FROM {TABLA} WHERE empresa=? AND periodo=?", (empresa, int(periodo)))
    con.commit()
