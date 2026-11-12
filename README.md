# Panel financiero · Área Limpia

Aplicación Streamlit con módulos de Presupuesto, Flujo de caja, Estados financieros, Cartera,
Combustible, Mantenimiento, Operaciones, Nómina y Administración.

## Ejecución
    pip install -r requirements.txt
    streamlit run app.py

## Configuración
- Base de datos: SQLite local (`presupuesto.db`) por defecto; PostgreSQL con `DATABASE_URL`
  (o `[base_datos] url` en secrets). Esquema opcional con `DB_SCHEMA`.
- Usuarios y permisos: `.streamlit/secrets.toml` con secciones `[usuarios]` y `[acceso]` (no se versiona).
