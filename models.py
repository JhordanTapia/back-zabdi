from sqlalchemy import Column, Integer, String, Float, Date, ForeignKey, Text, DateTime, Numeric, CheckConstraint, Boolean
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from database import Base
import uuid


# 0. USUARIOS (Simulación Supabase)
class Usuario(Base):
    __tablename__ = "usuarios"

    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    email = Column(String(255), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)

    perfil = relationship("Perfil", back_populates="usuario", uselist=False, cascade="all, delete-orphan")


# 1. PERFILES
class Perfil(Base):
    __tablename__ = "perfiles"

    id = Column(UUID(as_uuid=True), ForeignKey("usuarios.id", ondelete="CASCADE"), primary_key=True)
    nombre_completo = Column(String(150), nullable=False)
    rol = Column(String(20), nullable=False, default='Operador')
    creado_en = Column(DateTime(timezone=True), server_default=func.now())

    usuario = relationship("Usuario", back_populates="perfil")
    presupuestos_creados = relationship("Presupuesto", back_populates="creador")


# 2. PROVEEDORES (Aquí entraría "ASTILLEROS ANDESA")
class Proveedor(Base):
    __tablename__ = "proveedores"

    id = Column(Integer, primary_key=True, index=True)
    ruc = Column(String(11), unique=True, nullable=True)
    razon_social = Column(String(255), nullable=False)
    nombre_comercial = Column(String(255), nullable=True)

    # --- NUEVOS CAMPOS GEOGRÁFICOS ---
    direccion = Column(String(255), nullable=True)
    distrito = Column(String(100), nullable=True)
    provincia = Column(String(100), nullable=True)
    departamento = Column(String(100), nullable=True)

    contacto_nombre = Column(String(100))
    telefono = Column(String(20))
    email = Column(String(100))
    moneda_defecto = Column(String(3), default='PEN')
    creado_en = Column(DateTime(timezone=True), server_default=func.now())

    presupuestos = relationship("Presupuesto", back_populates="proveedor")


# 3. CATÁLOGO ESTÁNDAR
class CatalogoEstandar(Base):
    __tablename__ = "catalogo_estandar"

    id = Column(Integer, primary_key = True, index = True)
    codigo_estandar = Column(String(50), unique=True, nullable=False, index=True)
    descripcion = Column(Text, nullable=False)
    unidad_medida = Column(String(20), nullable=False)
    precio_referencial = Column(Numeric(12, 2), default=0.00)
    creado_en = Column(DateTime(timezone=True), server_default=func.now())

# Añade esta nueva clase arriba de tu clase Presupuesto
class EstadoPresupuesto(Base):
    __tablename__ = "estados_presupuesto"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(50), unique=True, nullable=False)
    color_hex = Column(String(20), default="bg-slate-500")

    presupuestos = relationship("Presupuesto", back_populates="estado_actual")
# 4. PRESUPUESTOS (La cabecera real)
class Presupuesto(Base):
    __tablename__ = "presupuestos"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    nro_presupuesto = Column(String(50))
    id_proveedor = Column(Integer, ForeignKey("proveedores.id", ondelete="SET NULL"), index=True)
    proyecto_embarcacion = Column(String(255), nullable=False, index=True)
    fecha_emision = Column(Date)

    # --- CAMBIO AQUÍ ---
    id_estado = Column(Integer, ForeignKey("estados_presupuesto.id"), default=1)
    # -------------------

    moneda = Column(String(3), default='PEN')
    url_pdf_storage = Column(Text)
    notas = Column(Text, nullable=True)
    eliminado = Column(Boolean, default=False)  # <--- NUEVO CAMPO PARA SOFT DELETE
    creado_por = Column(UUID(as_uuid=True), ForeignKey("perfiles.id"))
    creado_en = Column(DateTime(timezone=True), server_default=func.now())
    actualizado_en = Column(DateTime(timezone=True), server_default=func.now(), onupdate=func.now())

    # Relaciones
    proveedor = relationship("Proveedor", back_populates="presupuestos")
    creador = relationship("Perfil", back_populates="presupuestos_creados")
    items = relationship("ItemPresupuesto", back_populates="presupuesto", cascade="all, delete-orphan")
    estado_actual = relationship("EstadoPresupuesto", back_populates="presupuestos")  # <-- Nueva relación


# 5. DETALLE DEL PRESUPUESTO (Los items extraídos por la IA)
class ItemPresupuesto(Base):
    __tablename__ = "items_presupuesto"

    id = Column(Integer, primary_key = True, index = True, autoincrement = True)  # Tu BIGSERIAL
    id_presupuesto = Column(Integer, ForeignKey("presupuestos.id", ondelete="CASCADE"), index=True)
    descripcion_original = Column(Text, nullable=False)  # El "detalle_actividad" de la IA va aquí
    id_material_estandar = Column(Integer, ForeignKey("catalogo_estandar.id", ondelete="SET NULL"))
    cantidad = Column(Numeric(10, 2), nullable=False, default=1.00)
    precio_unitario = Column(Numeric(12, 2), nullable=False, default=0.00)
    horas_hombre = Column(Numeric(8, 2), default=0.00)
    notas = Column(Text, nullable=True)  # <--- NUEVA COLUMNA EN EL DETALLE

    creado_en = Column(DateTime(timezone=True), server_default=func.now())

    # Relaciones
    presupuesto = relationship("Presupuesto", back_populates="items")
    material_estandar = relationship("CatalogoEstandar")