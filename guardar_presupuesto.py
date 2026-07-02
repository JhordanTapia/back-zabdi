from models import Presupuesto, ItemPresupuesto, Perfil, Proveedor
from datetime import datetime

def guardar_datos_revisados_en_bd(data_ia, usuario_id, db):
    try:
        print("\n3. Mapeando el JSON revisado a las tablas reales...")

        # --- PARCHE: CREAR PERFIL SI NO EXISTE ---
        perfil = db.query(Perfil).filter(Perfil.id == usuario_id).first()
        if not perfil:
            print("   ! Este usuario no tiene perfil. Creando perfil por defecto...")
            nuevo_perfil = Perfil(id=usuario_id, nombre_completo="Usuario Operador", rol="Operador")
            db.add(nuevo_perfil)
            db.flush() # Lo guardamos temporalmente en esta transacción

        # --- PARSEAR FECHA ---
        fecha_parseada = None
        if data_ia.get("fecha") and data_ia.get("fecha") != "NO ENCONTRADO":
            try:
                fecha_parseada = datetime.strptime(data_ia["fecha"], "%Y-%m-%d").date()
            except:
                fecha_parseada = datetime.now().date()

        # --- NUEVO: BUSCAR AL PROVEEDOR (ASTILLERO) EN LA BD ---
        nombre_cliente_ia = data_ia.get("cliente", "")
        proveedor_id = None
        moneda_asignada = 'PEN'  # Moneda por defecto (Soles)

        if nombre_cliente_ia and nombre_cliente_ia != "NO ENCONTRADO":
            # 1. Filtramos palabras comunes para quedarnos con el núcleo del nombre
            palabras_comunes = ["ASTILLERO", "ASTILLEROS", "SAC", "EIRL", "SRL", "S.A.C.", "S.A.", "DE", "LA", "LOS",
                                "LAS"]
            palabras_reales = [p for p in nombre_cliente_ia.upper().split() if p not in palabras_comunes]

            # 2. Usamos la primera palabra clave encontrada (ej. "ANDESA"), si no, usamos todo
            termino_busqueda = palabras_reales[0] if palabras_reales else nombre_cliente_ia

            # 3. Buscamos en la BD usando ese término central
            proveedor_bd = db.query(Proveedor).filter(Proveedor.razon_social.ilike(f"%{termino_busqueda}%")).first()

            if proveedor_bd:
                proveedor_id = proveedor_bd.id
                moneda_asignada = proveedor_bd.moneda_defecto
                print(f"   ✓ Proveedor encontrado: {proveedor_bd.razon_social} (Moneda: {moneda_asignada})")
            else:
                print(f"   ! Proveedor no encontrado. (Se buscó la palabra clave: '{termino_busqueda}')")
            if proveedor_bd:
                proveedor_id = proveedor_bd.id
                moneda_asignada = proveedor_bd.moneda_defecto
                print(f"   ✓ Proveedor encontrado: {proveedor_bd.razon_social} (Moneda: {moneda_asignada})")
            else:
                texto_astillero = nombre_cliente_ia.strip().upper()
                nuevo_proveedor = Proveedor(
                    razon_social=texto_astillero,
                    nombre_comercial=texto_astillero
                )
                db.add(nuevo_proveedor)
                db.flush()
                proveedor_id = nuevo_proveedor.id
                print(f"   ! Proveedor nuevo creado como FANTASMA: '{texto_astillero}' (Pendiente de RUC)")

        # 1. Creamos la Cabecera (Tabla: presupuestos)
        nuevo_presupuesto = Presupuesto(
            nro_presupuesto=data_ia.get("numero_cotizacion") if data_ia.get(
                "numero_cotizacion") != "NO ENCONTRADO" else None,
            id_proveedor=proveedor_id,
            proyecto_embarcacion=data_ia.get("embarcacion", "Embarcación Desconocida"),
            fecha_emision=fecha_parseada,
            # estado="Borrador",
            moneda=data_ia.get("moneda", "PEN"),
            notas=data_ia.get("notas", None),
            creado_por=usuario_id
        )

        db.add(nuevo_presupuesto)
        db.flush() # Obtenemos el ID temporal

        print(f"   ✓ Cabecera creada con ID: {nuevo_presupuesto.id}")

        # 2. Creamos los Detalles (Tabla: items_presupuesto)
        # 2. Creamos los Detalles (Tabla: items_presupuesto)
        for item in data_ia.get("items", []):
            nuevo_item = ItemPresupuesto(
                id_presupuesto=nuevo_presupuesto.id,
                descripcion_original=item.get("detalle_actividad"),
                cantidad=item.get("cantidad", 1.0),
                precio_unitario=item.get("precio_unitario", 0.0),
                horas_hombre=0.0,
                notas=item.get("notas", None)  # <--- GUARDAMOS LA NOTA DEL ÍTEM
            )
            db.add(nuevo_item)

        # 3. Guardamos definitivamente en Postgres
        db.commit()
        return True, f"¡ÉXITO! Presupuesto '{nuevo_presupuesto.proyecto_embarcacion}' guardado."

    except Exception as e:
        db.rollback()
        print(f"Error al guardar en la BD: {e}")
        return False, str(e)