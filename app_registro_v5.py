import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import openpyxl
import os
import traceback

INGRESOS = {"Renta de Autos": 1, "Asesorías": 2, "Mecánica": 3,
            "Compra/Venta Automotriz": 4}
GASTOS = {"Mantenimiento": 1, "Legales": 2, "Planilla": 3,
          "Impuestos": 4, "Prestamo": 5}
NO_APLICA = "No aplica"

# ------------------------------------------------------------------
# Log de auditoría: deja constancia en un .txt junto al Excel de cada
# intento de registro (éxito o fallo), para poder diagnosticar casos
# donde el usuario cree que "no se guardó nada".
# ------------------------------------------------------------------
def log_path(path):
    base, _ = os.path.splitext(path)
    return base + "_registro_log.txt"

def log(path, msg):
    try:
        with open(log_path(path), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {msg}\n")
    except Exception:
        pass  # el log nunca debe impedir el registro


def sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws
    return None

def cols(ws):
    return {str(ws.cell(1,c).value).strip().lower(): c
            for c in range(1, ws.max_column+1)
            if ws.cell(1,c).value not in (None,"")}

def col(h, *names):
    for n in names:
        if n.lower() in h: return h[n.lower()]
    return None

def empty_row(ws, needed):
    r = 2
    while any(ws.cell(r,c).value not in (None,"") for c in needed): r += 1
    return r

def plates(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    result = []
    for name in ("Base de Datos Ingresos", "Base de Datos Gastos"):
        ws = sheet(wb,name)
        if ws:
            c = col(cols(ws),"Auto_ID")
            if c:
                for r in range(2,ws.max_row+1):
                    v=ws.cell(r,c).value
                    if v not in (None,""):
                        v=str(v).strip().upper()
                        if v and v not in result: result.append(v)
    wb.close()
    return sorted(result)

def copy_style(ws,r):
    if r <= 2: return
    for c in range(1,ws.max_column+1):
        if ws.cell(r-1,c).has_style: ws.cell(r,c)._style=ws.cell(r-1,c)._style

# ------------------------------------------------------------------
# Verificación post-guardado: vuelve a abrir el archivo (ya cerrado)
# en modo solo-lectura y confirma que los valores quedaron escritos
# en la fila esperada. Si algo no coincide, lanza un error claro en
# vez de dejar que el usuario crea que todo salió bien.
# ------------------------------------------------------------------
def verify_row(path, sheet_name, row, expected):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = sheet(wb, sheet_name)
        if ws is None:
            raise RuntimeError(f"No se encontró la hoja '{sheet_name}' al verificar.")
        for c, val in expected.items():
            actual = ws.cell(row, c).value
            if isinstance(val, datetime) and hasattr(actual, "date"):
                ok = actual.date() == val.date() if hasattr(actual, "date") else False
            elif isinstance(val, (int, float)) and isinstance(actual, (int, float)):
                ok = abs(actual - val) < 0.005
            else:
                ok = str(actual).strip() == str(val).strip()
            if not ok:
                raise RuntimeError(
                    f"Verificación falló en fila {row}, columna {c}: "
                    f"se esperaba {val!r} y se encontró {actual!r}."
                )
    finally:
        wb.close()

def save_income(path, date, iid, auto, amount):
    wb=openpyxl.load_workbook(path)
    ws=sheet(wb,"Base de Datos Ingresos")
    if not ws: raise ValueError("No existe 'Base de Datos Ingresos'.")
    h=cols(ws)
    cs=[col(h,"Fecha"),col(h,"Ingreso_ID"),col(h,"Auto_ID"),col(h,"Monto")]
    if any(x is None for x in cs): raise ValueError("Faltan columnas en ingresos.")
    r=empty_row(ws,cs)
    ws.cell(r,cs[0],date); ws.cell(r,cs[1],iid); ws.cell(r,cs[2],auto); ws.cell(r,cs[3],amount)
    ws.cell(r,cs[0]).number_format="dd/mm/yyyy"; ws.cell(r,cs[3]).number_format='#,##0.00'
    copy_style(ws,r)
    wb.save(path); wb.close()

    # Verificar que realmente quedó escrito antes de reportar éxito
    verify_row(path, "Base de Datos Ingresos", r,
               {cs[0]: date, cs[1]: iid, cs[2]: auto, cs[3]: amount})
    return r

def save_expense(path,date,eid,auto,parts,labor):
    wb=openpyxl.load_workbook(path)
    ws=sheet(wb,"Base de Datos Gastos")
    if not ws: raise ValueError("No existe 'Base de Datos Gastos'.")
    h=cols(ws)
    cs=[col(h,"Fecha"),col(h,"ID Gasto"),col(h,"Auto_ID"),
        col(h,"Costo Repuesto"),col(h,"Mano de obra"),col(h,"Total Gasto")]
    if any(x is None for x in cs): raise ValueError("Faltan columnas en gastos.")
    r=empty_row(ws,cs[:5]); copy_style(ws,r)
    ws.cell(r,cs[0],date); ws.cell(r,cs[1],eid); ws.cell(r,cs[2],auto)
    ws.cell(r,cs[3],parts); ws.cell(r,cs[4],labor)
    ws.cell(r,cs[5],f"={ws.cell(r,cs[3]).coordinate}+{ws.cell(r,cs[4]).coordinate}")
    ws.cell(r,cs[0]).number_format="dd/mm/yyyy"
    for c in cs[3:]: ws.cell(r,c).number_format='#,##0.00'
    wb.save(path); wb.close()

    # Verificar (la columna de Total Gasto es fórmula, no se compara su valor
    # porque openpyxl no la recalcula hasta que se abre en Excel)
    verify_row(path, "Base de Datos Gastos", r,
               {cs[0]: date, cs[1]: eid, cs[2]: auto, cs[3]: parts, cs[4]: labor})
    return r


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Registro de Ingresos y Gastos"); self.geometry("700x790")
        self.minsize(620,700)
        self.file=tk.StringVar(); self.kind=tk.StringVar(value="Ingreso")
        self.date=tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.auto=tk.StringVar(); self.category=tk.StringVar()
        self.amount=tk.StringVar(); self.parts=tk.StringVar(); self.labor=tk.StringVar()
        self.build(); self.bind_vars(); self.form(); self.validate()

    def report_callback_exception(self, exc, val, tb):
        # CLAVE: por defecto, Tkinter atrapa cualquier error ocurrido dentro
        # de un comando de botón (o cualquier callback) y solo lo imprime en
        # la consola — nunca muestra un cuadro de diálogo. Si la app se
        # ejecuta sin consola visible (doble clic, o un .exe sin consola),
        # ese error desaparece por completo y parece que "no pasó nada" al
        # presionar el botón. Este método sobreescribe ese comportamiento
        # para que CUALQUIER error, venga de donde venga, se muestre siempre.
        detalle = "".join(traceback.format_exception(exc, val, tb))
        try:
            log(self.file.get() or "sin_archivo", f"EXCEPCION NO CAPTURADA:\n{detalle}")
        except Exception:
            pass
        messagebox.showerror(
            "Error inesperado",
            f"Ocurrió un error que impidió completar la acción:\n\n{val}\n\n"
            f"(Se guardó el detalle técnico en el log de registro)."
        )

    def build(self):
        root=ttk.Frame(self,padding=20); root.pack(fill="both",expand=True)
        ttk.Label(root,text="Registro de Ingresos y Gastos",
                  font=("Segoe UI",18,"bold")).pack(anchor="w",pady=(0,15))

        f=ttk.LabelFrame(root,text="Archivo Excel",padding=10); f.pack(fill="x",pady=6)
        ttk.Entry(f,textvariable=self.file,state="readonly").pack(side="left",fill="x",expand=True,padx=5)
        ttk.Button(f,text="Seleccionar...",command=self.select).pack(side="left")
        ttk.Button(f,text="Probar escritura",command=self.test_write).pack(side="left",padx=(6,0))

        f=ttk.LabelFrame(root,text="1. Tipo de movimiento",padding=10); f.pack(fill="x",pady=6)
        ttk.Radiobutton(f,text="Ingreso",variable=self.kind,value="Ingreso",command=self.form).pack(side="left",padx=20)
        ttk.Radiobutton(f,text="Gasto",variable=self.kind,value="Gasto",command=self.form).pack(side="left")

        f=ttk.LabelFrame(root,text="2. Datos del movimiento",padding=12); f.pack(fill="x",pady=6)
        ttk.Label(f,text="Fecha (DD/MM/AAAA):").grid(row=0,column=0,sticky="w",pady=6)
        self.date_e=ttk.Entry(f,textvariable=self.date); self.date_e.grid(row=0,column=1,sticky="ew")
        ttk.Label(f,text="Ejemplo: 05/09/2026").grid(row=0,column=2,padx=8)
        ttk.Label(f,text="Auto_ID:").grid(row=1,column=0,sticky="w",pady=6)
        self.auto_c=ttk.Combobox(f,textvariable=self.auto,state="readonly"); self.auto_c.grid(row=1,column=1,sticky="ew")
        ttk.Button(f,text="Actualizar placas",command=self.refresh).grid(row=1,column=2,padx=8)
        ttk.Label(f,text="Categoría:").grid(row=2,column=0,sticky="w",pady=6)
        self.cat_c=ttk.Combobox(f,textvariable=self.category,state="readonly"); self.cat_c.grid(row=2,column=1,columnspan=2,sticky="ew")
        self.al=ttk.Label(f,text="Monto:")
        self.ae=ttk.Entry(f,textvariable=self.amount)
        self.pl=ttk.Label(f,text="Costo del repuesto:"); self.pe=ttk.Entry(f,textvariable=self.parts)
        self.ll=ttk.Label(f,text="Mano de obra:"); self.le=ttk.Entry(f,textvariable=self.labor)
        self.total=ttk.Label(f,text="Total del gasto: 0.00",font=("Segoe UI",11,"bold"))
        f.columnconfigure(1,weight=1)

        self.status=ttk.Label(root,text="Complete todos los campos correctamente.",foreground="red")
        self.status.pack(anchor="w",pady=6)

        # Banner de confirmación persistente (además del messagebox), para que
        # quede visible en pantalla cuál fue el último registro exitoso.
        self.confirm=ttk.Label(root,text="",foreground="#1a7d1a",font=("Segoe UI",10,"bold"))
        self.confirm.pack(anchor="w",pady=(0,4))

        self.reg=ttk.Button(root,text="Registrar movimiento",command=self.register,state="disabled")
        self.reg.pack(fill="x",pady=5)
        ttk.Button(root,text="Limpiar",command=self.clear).pack(fill="x",pady=3)

        h=ttk.LabelFrame(root,text="Últimos movimientos registrados",padding=8); h.pack(fill="both",expand=True,pady=8)
        cs=("fecha","tipo","categoria","auto","monto","fila")
        self.tree=ttk.Treeview(h,columns=cs,show="headings",height=7)
        for c,t,w in zip(cs,("Fecha","Tipo","Categoría","Auto_ID","Monto / Total","Fila Excel"),(85,70,170,85,105,80)):
            self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor="center")
        sb=ttk.Scrollbar(h,orient="vertical",command=self.tree.yview); self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")

    def bind_vars(self):
        for v in (self.file,self.kind,self.date,self.auto,self.category,self.amount,self.parts,self.labor):
            v.trace_add("write",lambda *x:self.validate())

    def form(self):
        if self.kind.get()=="Ingreso":
            self.cat_c["values"]=list(INGRESOS)
            self.al.grid(row=3,column=0,sticky="w",pady=6); self.ae.grid(row=3,column=1,columnspan=2,sticky="ew")
            for w in (self.pl,self.pe,self.ll,self.le,self.total): w.grid_remove()
        else:
            self.cat_c["values"]=list(GASTOS)
            self.al.grid_remove(); self.ae.grid_remove()
            self.pl.grid(row=3,column=0,sticky="w",pady=6); self.pe.grid(row=3,column=1,columnspan=2,sticky="ew")
            self.ll.grid(row=4,column=0,sticky="w",pady=6); self.le.grid(row=4,column=1,columnspan=2,sticky="ew")
            self.total.grid(row=5,column=0,columnspan=3,sticky="w",pady=8)
        self.category.set(""); self.amount.set(""); self.parts.set(""); self.labor.set(""); self.validate()

    def valid_date(self):
        s=self.date.get().strip()
        if len(s)!=10 or s[2]!="/" or s[5]!="/": return False
        if not (s[:2].isdigit() and s[3:5].isdigit() and s[6:].isdigit()): return False
        try:
            datetime(int(s[6:]),int(s[3:5]),int(s[:2])); return True
        except ValueError: return False

    def valid_money(self,s):
        s=s.strip().replace(",",".")
        if not s or s.count(".")>1: return False
        p=s.split(".")
        return p[0].isdigit() and (len(p)==1 or (p[1].isdigit() and len(p[1])<=2))

    def validate(self):
        if self.kind.get()=="Gasto":
            try:
                p=float(self.parts.get().replace(",",".")) if self.valid_money(self.parts.get()) else 0
                l=float(self.labor.get().replace(",",".")) if self.valid_money(self.labor.get()) else 0
                self.total.config(text=f"Total del gasto: {p+l:,.2f}")
            except: self.total.config(text="Total del gasto: 0.00")
        ok=bool(self.file.get() and self.valid_date() and self.auto.get() and self.category.get() and
                (self.valid_money(self.amount.get()) if self.kind.get()=="Ingreso" else
                 self.valid_money(self.parts.get()) and self.valid_money(self.labor.get())))
        self.reg.config(state="normal" if ok else "disabled")
        self.status.config(text="✓ Todos los campos son válidos." if ok else "Complete todos los campos correctamente.",
                           foreground="green" if ok else "red")

    def select(self):
        p=filedialog.askopenfilename(title="Selecciona la plantilla Excel",filetypes=[("Excel","*.xlsx")])
        if not p:return
        try:
            if not os.access(p, os.W_OK):
                raise PermissionError("El archivo es de solo lectura o no tienes permisos de escritura sobre él.")
            wb=openpyxl.load_workbook(p,read_only=True,data_only=True)
            i=sheet(wb,"Base de Datos Ingresos"); g=sheet(wb,"Base de Datos Gastos")
            if not i or not g: raise ValueError("El Excel debe tener las hojas Base de Datos Ingresos y Base de Datos Gastos.")
            hi, hg=cols(i),cols(g); wb.close()
            if any(col(hi,x) is None for x in ("Fecha","Ingreso_ID","Auto_ID","Monto")): raise ValueError("Faltan columnas en ingresos.")
            if any(col(hg,x) is None for x in ("Fecha","ID Gasto","Auto_ID","Costo Repuesto","Mano de obra","Total Gasto")): raise ValueError("Faltan columnas en gastos.")
            self.file.set(p); self.confirm.config(text=""); self.refresh()
        except PermissionError as e:
            messagebox.showerror("Archivo bloqueado o sin permisos",
                                  f"{e}\n\nCierra el archivo en Excel (o cualquier otro programa) y vuelve a intentarlo.")
        except Exception as e: messagebox.showerror("Plantilla inválida",str(e))

    def test_write(self):
        # Prueba independiente del formulario: escribe un valor de marca en una
        # celda muy alejada de los datos reales (columna Z, fila 1000 de la
        # hoja de Ingresos), lo verifica, y lo borra. Sirve para confirmar de
        # forma aislada si el problema es de acceso al archivo o del formulario.
        path=self.file.get()
        if not path:
            messagebox.showwarning("Falta el archivo","Primero selecciona el archivo Excel."); return
        marca=f"prueba_{datetime.now().strftime('%H%M%S')}"
        try:
            wb=openpyxl.load_workbook(path)
            ws=sheet(wb,"Base de Datos Ingresos")
            ws.cell(1000,26,marca)  # columna Z, muy lejos de tus datos reales
            wb.save(path); wb.close()
            wb2=openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws2=sheet(wb2,"Base de Datos Ingresos")
            leido=ws2.cell(1000,26).value
            wb2.close()
            if leido != marca:
                raise RuntimeError(f"Se escribió '{marca}' pero al releer se encontró '{leido}'.")
            # limpiar la celda de prueba
            wb3=openpyxl.load_workbook(path)
            sheet(wb3,"Base de Datos Ingresos").cell(1000,26,None)
            wb3.save(path); wb3.close()
            messagebox.showinfo("Prueba exitosa",
                "✓ Se pudo escribir, leer y confirmar correctamente en el archivo.\n\n"
                "Si el registro normal sigue sin funcionar, el problema está en el "
                "formulario, no en el acceso al Excel.")
        except PermissionError:
            messagebox.showerror("Prueba fallida",
                "No se pudo escribir: el archivo está abierto en Excel u otro programa, "
                "o no tienes permisos sobre él.")
        except Exception as e:
            messagebox.showerror("Prueba fallida", f"No se pudo escribir ni verificar:\n\n{e}")

    def refresh(self):
        if not self.file.get(): return
        try:
            opts=[NO_APLICA]+plates(self.file.get()); self.auto_c["values"]=opts
            if self.auto.get() not in opts:self.auto.set("")
            self.validate()
        except Exception as e: messagebox.showerror("Error",str(e))

    def register(self):
        path=self.file.get()
        log(path or "sin_archivo", "Botón 'Registrar movimiento' presionado.")
        if not path:
            messagebox.showwarning("Falta el archivo","Primero selecciona el archivo Excel.")
            return
        if str(self.reg["state"]) != "normal":
            messagebox.showwarning("Faltan datos",
                "Completa todos los campos correctamente antes de registrar.\n"
                "Revisa el mensaje en rojo bajo el formulario.")
            return
        try:
            d=datetime.strptime(self.date.get(),"%d/%m/%Y"); auto=self.auto.get(); cat=self.category.get()

            # Reintenta automáticamente si el archivo está bloqueado (p.ej. abierto
            # en Excel), en vez de fallar una sola vez y dejar al usuario sin saber
            # si realmente se guardó o no.
            while True:
                try:
                    if self.kind.get()=="Ingreso":
                        amount=float(self.amount.get().replace(",",".")); iid=INGRESOS[cat]
                        row=save_income(path,d,iid,auto,amount); total=amount
                        msg=(f"Ingreso registrado correctamente.\n\n"
                             f"Hoja: Base de Datos Ingresos\nFila: {row}\n"
                             f"ID: {iid}\nAuto_ID: {auto}\nMonto: {amount:,.2f}")
                    else:
                        p=float(self.parts.get().replace(",",".")); l=float(self.labor.get().replace(",",".")); total=p+l; eid=GASTOS[cat]
                        row=save_expense(path,d,eid,auto,p,l)
                        msg=(f"Gasto registrado correctamente.\n\n"
                             f"Hoja: Base de Datos Gastos\nFila: {row}\n"
                             f"ID: {eid}\nAuto_ID: {auto}\nRepuesto: {p:,.2f}\n"
                             f"Mano de obra: {l:,.2f}\nTotal: {total:,.2f}")
                    break
                except PermissionError:
                    log(path, f"PermissionError al intentar guardar ({self.kind.get()}, auto={auto}).")
                    if not messagebox.askretrycancel(
                        "Archivo bloqueado",
                        "No se pudo guardar porque el Excel está abierto en otro programa "
                        "(por ejemplo, Excel o OneDrive sincronizando).\n\n"
                        "Cierra el archivo y presiona 'Reintentar', o 'Cancelar' para abortar el registro."):
                        return

            log(path, f"OK: {self.kind.get()} fila {row} auto={auto} categoria={cat} total={total:,.2f}")

            self.tree.insert("",0,values=(d.strftime("%d/%m/%Y"),self.kind.get(),cat,auto,f"{total:,.2f}",row))
            for item in self.tree.get_children()[10:]: self.tree.delete(item)

            # Confirmación visible en pantalla (además del cuadro emergente)
            self.confirm.config(
                text=f"✓ Registrado correctamente en la fila {row} de "
                     f"'{'Base de Datos Ingresos' if self.kind.get()=='Ingreso' else 'Base de Datos Gastos'}'."
            )
            messagebox.showinfo("Registro exitoso",msg)
            self.clear(keep_file=True)

        except RuntimeError as e:
            # Falla de verificación: el guardado se ejecutó pero al releer el
            # archivo los datos no coinciden con lo esperado.
            log(path, f"FALLO DE VERIFICACIÓN: {e}")
            messagebox.showerror(
                "No se pudo confirmar el registro",
                f"El archivo se guardó, pero al verificarlo los datos no coinciden con lo esperado.\n\n"
                f"Detalle: {e}\n\nRevisa el archivo manualmente antes de continuar."
            )
        except Exception as e:
            log(path, f"ERROR al registrar: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error al registrar",str(e))

    def clear(self,keep_file=False):
        if not keep_file:self.file.set("")
        self.date.set(datetime.now().strftime("%d/%m/%Y")); self.auto.set(""); self.category.set("")
        self.amount.set(""); self.parts.set(""); self.labor.set(""); self.validate()

if __name__=="__main__":
    App().mainloop()
