# -*- coding: utf-8 -*-
"""
Sistema de Bodega de Emergencias - Hospital Nacional Dr. Luis Edmundo Vásquez, Chalatenango.

Programa de escritorio para Windows (Python + Tkinter, empaquetado con PyInstaller).
No necesita internet. Los datos se guardan como archivos JSON (texto) en
Documentos\\BodegaEmergencias:

    datos\\AAAA-MM_SERVICIO.json      uno por cada mes y servicio
    modelos\\SERVICIO.json            lista modelo de insumos de cada servicio
    matrices\\Matriz_... .json        matrices por período
    documentos\\AAAA-MM SERVICIO\\    documentos PDF generados
    config.json                       encabezado y numeraciones
"""
import json, os, re, sys, shutil, unicodedata, uuid, datetime, subprocess, zipfile
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

VERSION = '1.0 (27/11/2026)'
APP = 'Sistema de Bodega de Emergencias'
MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
DSEM = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
CFG_DEF = {
    'l1': 'M03- ATENCIÓN EN SALUD INTEGRAL E INTEGRADA A LA PERSONA EN EL CURSO DE VIDA CON ENFOQUE DE APS',
    'l2': 'M03-PAIE PROPORCIONAR ATENCION DE SALUD INTREGRALES EN EMERGENCIA',
    'tit': 'MATRIZ DE ENTRADAS Y SALIDAS DE INSUMOS', 'cod': 'M03-PAIE-UE-MTZ-01', 'ver': 'Versión 01',
    'vig': 'Vigencia 15/10/2026', 'resp': '', 'hospital': 'HOSPITAL NACIONAL Dr. LUIS EDMUNDO VASQUEZ, CHALATENANGO',
    'f1': 'Elaborado por', 'f2': 'Revisado por', 'f3': 'Vo. Bo. Jefatura', 'firmas': False, 'folios': {}, 'recibos': {},
}


# --------------------------------------------------------------------------- utilidades
def recurso(nombre):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)


def carpeta_documentos():
    if os.name == 'nt':
        try:
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # CSIDL_PERSONAL = Documentos
            if buf.value:
                return buf.value
        except Exception:
            pass
    return os.path.join(os.path.expanduser('~'), 'Documents')


def nr(s):
    s = unicodedata.normalize('NFD', str(s or ''))
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn').upper()
    return re.sub(r'[^A-Z0-9/#]', '', s)


ALIAS = {
    'C10601290': 'C10602095',  # tuberculina anotada en aguja de pluma
    'C10501030': 'N' + nr('MASCARILLA DE REINHALACION RESERVORIO PARA ADMINISTRACION DE OXIGENO Y TUBO EXTENSION DE 2,1m (7 PIES),TAMAÑO INFANTIL, DESCARTABLE'),
    'N' + nr('ESTOQUINETE 2'): 'N' + nr('ESTOQUINETE DE 2 PULGADAS'),
    'N' + nr('ESTOQUINETE 4'): 'N' + nr('ESTOQUINETE DE 4 PULGADAS'),
    'N' + nr('CATETER DE VENA CENTRAL'): 'C10605267',
}


def clave(it):
    k = 'C' + str(it.get('c')).strip() if it.get('c') else 'N' + nr(it.get('n'))
    return ALIAS.get(k, k)


def nuevo_id(p='i'):
    return p + uuid.uuid4().hex[:12]


def fmt(n):
    if n is None or n == '':
        return ''
    try:
        n = float(n)
    except Exception:
        return str(n)
    if n == int(n):
        return f'{int(n):,}'
    return f'{n:,.2f}'


def fecha_txt(f):
    return '/'.join(reversed(f.split('-'))) if f else ''


def vto_txt(v):
    return f'{v[5:7]}/{v[:4]}' if v and len(v) >= 7 else ''


def nombre_mes(mid):
    y, m = mid.split('-')
    return f'{MESES[int(m) - 1]} {y}'


def dias_mes(mid):
    y, m = map(int, mid.split('-'))
    n = datetime.date(y + (m == 12), m % 12 + 1, 1) - datetime.timedelta(days=1)
    return n.day


def mes_siguiente(mid):
    y, m = map(int, mid.split('-'))
    return f'{y + (m == 12)}-{m % 12 + 1:02d}'


def num(v):
    if v is None or v == '':
        return None
    try:
        return float(v) if '.' in str(v) else int(v)
    except Exception:
        return None


def eval_cant(s):
    s = re.sub(r'\s', '', str(s or ''))
    if not s:
        return 0
    if not re.fullmatch(r'\d+(\+\d+)*', s):
        raise ValueError('Cantidad no válida: ' + s)
    return sum(int(x) for x in s.split('+'))


def abrir(ruta):
    try:
        if os.name == 'nt':
            os.startfile(ruta)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', ruta])
        else:
            subprocess.Popen(['xdg-open', ruta])
    except Exception:
        pass


# --------------------------------------------------------------------------- datos
class Datos:
    def __init__(self, base=None):
        self.base = base or os.path.join(carpeta_documentos(), 'BodegaEmergencias')
        for d in ('datos', 'modelos', 'matrices', 'documentos', 'respaldos'):
            os.makedirs(os.path.join(self.base, d), exist_ok=True)
        self.cache = {}
        self.cfg = dict(CFG_DEF)
        p = os.path.join(self.base, 'config.json')
        if os.path.exists(p):
            try:
                self.cfg.update(json.load(open(p, encoding='utf-8')))
            except Exception:
                pass

    # --- archivos
    def _escribir(self, ruta, obj):
        tmp = ruta + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        os.replace(tmp, ruta)

    def guardar_cfg(self):
        self._escribir(os.path.join(self.base, 'config.json'), self.cfg)

    def ruta_mes(self, serv, mid):
        return os.path.join(self.base, 'datos', f'{mid}_{serv}.json')

    def servicios(self):
        s = set()
        for f in os.listdir(os.path.join(self.base, 'datos')):
            m = re.match(r'(\d{4}-\d{2})_(.+)\.json$', f)
            if m:
                s.add(m.group(2))
        for f in os.listdir(os.path.join(self.base, 'modelos')):
            if f.endswith('.json'):
                s.add(f[:-5])
        return sorted(s) or ['EMERGENCIA']

    def meses(self, serv):
        r = []
        for f in os.listdir(os.path.join(self.base, 'datos')):
            m = re.match(r'(\d{4}-\d{2})_(.+)\.json$', f)
            if m and m.group(2) == serv:
                r.append(m.group(1))
        return sorted(r)

    def mes(self, serv, mid):
        k = (serv, mid)
        if k not in self.cache:
            p = self.ruta_mes(serv, mid)
            if os.path.exists(p):
                self.cache[k] = json.load(open(p, encoding='utf-8'))
            else:
                return None
        d = self.cache[k]
        for kk, v in (('items', []), ('vales', []), ('retiros', []), ('conteo', {}), ('antisoborno', {})):
            d.setdefault(kk, v)
        return d

    def guardar_mes(self, serv, mid):
        d = self.cache.get((serv, mid))
        if d is not None:
            d['guardado'] = datetime.datetime.now().isoformat(timespec='seconds')
            self._escribir(self.ruta_mes(serv, mid), d)

    def crear_mes(self, serv, mid, items):
        d = {'version': 1, 'servicio': serv, 'mes': mid, 'nombre': nombre_mes(mid), 'items': items,
             'vales': [], 'retiros': [], 'conteo': {}, 'antisoborno': {}}
        self.cache[(serv, mid)] = d
        self.guardar_mes(serv, mid)
        return d

    def modelo(self, serv):
        p = os.path.join(self.base, 'modelos', serv + '.json')
        if os.path.exists(p):
            return json.load(open(p, encoding='utf-8'))
        return {'servicio': serv, 'items': []}

    def guardar_modelo(self, serv, mod):
        self._escribir(os.path.join(self.base, 'modelos', serv + '.json'), mod)

    # --- cálculos: la existencia pasa sola de un mes al siguiente
    def calcular(self, serv, hasta=None):
        res, prev, inc = {}, {}, {}
        for mid in self.meses(serv):
            m = self.mes(serv, mid)
            r, vistos, nxt, ninc = {}, {}, dict(prev), dict(inc)
            for it in m['items']:
                if it.get('t') == 'i':
                    r[it['id']] = {'rec': 0, 'fechas': [], 'vto': '', 'cons': sum(x or 0 for x in (it.get('d') or []) if isinstance(x, (int, float)))}
            for v in m['vales']:
                for l in v['lineas']:
                    x = r.get(l['item'])
                    if not x:
                        continue
                    x['rec'] += l.get('cant') or 0
                    if v.get('fecha') and v['fecha'] not in x['fechas']:
                        x['fechas'].append(v['fecha'])
                    if l.get('vto') and (not x['vto'] or l['vto'] < x['vto']):
                        x['vto'] = l['vto']
            for it in m['items']:
                if it.get('t') != 'i':
                    continue
                x = r[it['id']]
                x['fechas'].sort()
                k = clave(it)
                vistos[k] = vistos.get(k, 0) + 1
                if vistos[k] > 1:
                    k += '#' + str(vistos[k])
                man = num(it.get('ini'))
                antraw = prev.get(k)
                ant = 0 if (antraw is not None and antraw < 0) else antraw
                x.update(clave=k, ant=ant, antNeg=antraw if (antraw is not None and antraw < 0) else None, iniManual=man is not None)
                x['ini'] = man if man is not None else ant
                x['iniAuto'] = man is None and ant is not None
                x['ajuste'] = (man - ant) if (man is not None and ant is not None and man != ant) else None
                x['sinDatos'] = x['ini'] is None and x['rec'] == 0 and x['cons'] == 0
                x['sal'] = (x['ini'] or 0) + x['rec'] - x['cons']
                x['incompleto'] = x['sal'] < 0 or (man is None and ((antraw is not None and antraw < 0) or inc.get(k)))
                if not x['sinDatos']:
                    nxt[k] = x['sal']
                    ninc[k] = x['incompleto']
            res[mid] = r
            prev, inc = nxt, ninc
            if mid == hasta:
                break
        return res

    def calc_mes(self, serv, mid):
        return self.calcular(serv, mid).get(mid, {})

    # --- numeraciones
    def folio_hoja(self, serv, fecha, asignar=True):
        F = self.cfg.setdefault('folios', {})
        k = serv + '|' + fecha
        if k in F or not asignar:
            return F.get(k)
        mx = max([0] + [int(v) for v in F.values()])
        for mid in self.meses(serv):
            for r in self.mes(serv, mid)['retiros']:
                try:
                    mx = max(mx, int(r.get('folio') or 0))
                except Exception:
                    pass
        F[k] = mx + 1
        self.guardar_cfg()
        return F[k]

    def siguiente_recibo(self, anio):
        R = self.cfg.setdefault('recibos', {})
        R[str(anio)] = int(R.get(str(anio), 0)) + 1
        self.guardar_cfg()
        return f'{R[str(anio)]:03d}/{anio}'

    # --- importar el respaldo del programa anterior (navegador)
    def importar_respaldo(self, obj, serv='EMERGENCIA'):
        S = obj.get('datos', obj) if isinstance(obj, dict) else None
        if not S or 'meses' not in S:
            raise ValueError('El archivo no es un respaldo del Sistema de Bodega.')
        n = 0
        for m in S['meses']:
            mid = m['id']
            items = []
            for it in m['items']:
                o = {'id': it['id'], 't': it.get('t', 'i'), 'n': it.get('n', '')}
                if it.get('c'):
                    o['c'] = it['c']
                if o['t'] == 'i':
                    if it.get('ini') not in (None, ''):
                        o['ini'] = it['ini']
                    o['d'] = list(it.get('d') or [])
                items.append(o)
            d = self.crear_mes(serv, mid, items)
            d['nombre'] = m.get('nombre') or nombre_mes(mid)
            d['vales'] = [{'id': v['id'], 'num': v.get('num', ''), 'fecha': v.get('fecha', ''), 'rec': v.get('rec', ''),
                           'lineas': [{'item': l['item'], 'cant': l.get('cant', 0), 'vto': l.get('vto', ''), 'obs': l.get('obs', '')} for l in v.get('lineas', [])]}
                          for v in S.get('vales', []) if v.get('mes') == mid]
            d['retiros'] = [r for r in S.get('retiros', []) if r.get('mes') == mid]
            cu = (S.get('cuadres') or {}).get(mid) or {}
            d['conteo'] = {iid: k.get('f') for iid, k in (cu.get('c') or {}).items() if k.get('f') is not None}
            d['antisoborno'] = (S.get('antisob') or {}).get(mid) or {}
            self.guardar_mes(serv, mid)
            n += 1
        cfg = S.get('cfg') or {}
        for k in ('l1', 'l2', 'tit', 'cod', 'ver', 'vig', 'resp', 'f1', 'f2', 'f3'):
            if cfg.get(k):
                self.cfg[k] = cfg[k]
        self.guardar_cfg()
        if S['meses']:
            ult = S['meses'][-1]
            self.guardar_modelo(serv, {'servicio': serv, 'items': [{'t': it.get('t', 'i'), 'n': it.get('n', ''), 'c': it.get('c', '')} for it in ult['items']]})
        return n

    def respaldo_zip(self):
        destino = os.path.join(self.base, 'respaldos', 'Respaldo ' + datetime.datetime.now().strftime('%Y-%m-%d %Hh%M') + '.zip')
        with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as z:
            for sub in ('datos', 'modelos', 'matrices'):
                for f in os.listdir(os.path.join(self.base, sub)):
                    z.write(os.path.join(self.base, sub, f), sub + '/' + f)
            z.write(os.path.join(self.base, 'config.json') if os.path.exists(os.path.join(self.base, 'config.json')) else recurso('logo.png'), 'config.json')
        return destino


# --------------------------------------------------------------------------- nombres cortos para la hoja de pedido
NC = {'10602095': 'Jeringa de BCG y tuberculina 1 ml', '10602005': 'Jeringa de 3 ML (aguja 21 x 1½)', '10602040': 'Jeringa de 5 ML', '10602050': 'Jeringa de 10 ML',
      '10602060': 'Jeringa de 20 ML', '10602075': 'Jeringa de 50-60 ML', '10605165': 'Catéter intravenoso #18', '10605180': 'Catéter intravenoso #20', '10605190': 'Catéter intravenoso #22',
      '10605200': 'Catéter intravenoso #24', '12100165': 'Sellos de heparina', '10600500': 'Válvula de 3 vías', '10600250': 'Micro gotero descartable', '10600110': 'Descartables para suero sin aguja',
      '10701005': 'Adhesivo micro poroso 1/2', '10701010': 'Adhesivo micro poroso 1x10', '10701015': 'Adhesivo micro poroso 2x10', '10701020': 'Adhesivo micro poroso 3x10', '10701025': 'Adhesivo micro poroso 4x10',
      '10701030': 'Adhesivo tela seda 1x10', '10701035': 'Adhesivo tela seda 2x10', '10701040': 'Adhesivo tela seda 3x10', '10705025': 'Venda de gasa 2x10', '10705030': 'Venda de gasa 3x10',
      '10705035': 'Venda de gasa 4x10', '10705065': 'Venda elástica 4"', '10705085': 'Venda enyesada 4"', '10705090': 'Venda enyesada 6"', '10705115': 'Fibra de vidrio 2"', '10705118': 'Fibra de vidrio 3"',
      '10705120': 'Fibra de vidrio 4"', '10705121': 'Fibra de vidrio 5"', '10601005': 'Aguja descartable 18', '10601020': 'Aguja descartable 21', '10605267': 'Catéter de vena central',
      '10605540': 'Tubo de tórax 34', '10605550': 'Tubo de tórax 36', '10610165': 'Sonda Foley #14', '10610175': 'Sonda Foley #18', '10610180': 'Sonda Foley #20', '10600040': 'Bolsa recolectora de orina adulto',
      '10610110': 'Sonda nasogástrica #14', '10610115': 'Sonda nasogástrica #16', '10606010': 'Guante quirúrgico 6½', '10606015': 'Guante quirúrgico 7', '10606020': 'Guante quirúrgico 7½',
      '10606025': 'Guante quirúrgico 8', '10606040': 'Guantes examen talla M', '10606035': 'Guantes examen talla L', '10704010': 'Huata quirúrgica 4"', '10704015': 'Huata quirúrgica 6"',
      '10500020': 'Frasco humidificador de oxígeno', '10500040': 'Mascarilla para nebulizar adulto', '10500045': 'Mascarilla para nebulizar pediátrica', '10604005': 'Bigotera adulto',
      '10604010': 'Bigotera pediátrica', '10604015': 'Bigotera neonatal', '10501055': 'Mascarilla de reservorio adulto', '11103085': 'Nylon 3/0', '11103110': 'Nylon 2/0', '11800010': 'Alcohol litro',
      '11800019': 'Alcohol gel', '11800060': 'Clorhexidina', '11800065': 'Jabón yodado', '11800075': 'Jabón para manos (galón)', '11700045': 'Jalea lubricante (CAYE)', '11700030': 'Frasco lubricante gel / electro gel',
      '12100035': 'Brazalete adulto', '12100040': 'Brazalete pediátrico', '10608045': 'Mascarilla quirúrgica adulto', '10900019': 'Mascarilla N95 (KN95)', '10600235': 'Gorro para médico',
      '10200023': 'Papel para monitor fetal (NST)', '10600030': 'Baja lenguas de madera', '10609005': 'Condones', '10702105': 'Apósito transparente (Tegaderm)', '10202005': 'Electrodos adulto',
      '12100168': 'Termómetro digital'}


def nombre_hoja(it):
    if it.get('nc'):
        return it['nc']
    k = clave(it)
    if k.startswith('C') and k[1:] in NC:
        return NC[k[1:]]
    s = str(it.get('n', '')).split(',')[0]
    s = re.sub(r'\s+', ' ', re.sub(r'(?i)MONOFILAMENTO |PARA ADMINISTRACION DE OXIGENO| CON AGUJA.*$', '', s)).strip()
    s = s[:1] + s[1:].lower()
    return s if len(s) <= 36 else s[:35] + '…'


# --------------------------------------------------------------------------- documentos PDF (reportlab)
def _rl():
    from reportlab.lib.pagesizes import letter, legal, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfgen import canvas as rlcanvas
    return locals()


class PDF:
    def __init__(self, app):
        self.app = app
        self.D = app.D

    def carpeta(self, serv, mid):
        p = os.path.join(self.D.base, 'documentos', f'{mid} {serv}')
        os.makedirs(p, exist_ok=True)
        return p

    def _num_canvas(self):
        R = _rl()

        class NumCanvas(R['rlcanvas'].Canvas):
            def __init__(s, *a, **k):
                super().__init__(*a, **k)
                s._pags = []

            def showPage(s):
                s._pags.append(dict(s.__dict__))
                s._startPage()

            def save(s):
                n = len(s._pags)
                for st in s._pags:
                    s.__dict__.update(st)
                    s.setFont('Helvetica', 7.5)
                    s.drawCentredString(s._pagesize[0] / 2, 16, f'Página {s._pageNumber} de {n}')
                    super().showPage()
                super().save()
        return NumCanvas

    def encabezado(self, c, ancho, alto, titulo, serv, mid, margen=26):
        """Encabezado M03 dibujado en cada página."""
        cfg = self.D.cfg
        y = alto - margen
        h = 16
        x0, w_logo, w_der = margen, 110, 150
        w_mid = ancho - 2 * margen - w_logo - w_der
        c.setLineWidth(0.8)
        c.rect(x0, y - 3 * h, w_logo, 3 * h)
        try:
            c.drawImage(recurso('logo.png'), x0 + 6, y - 3 * h + 4, w_logo - 12, 3 * h - 8, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
        filas = [(cfg['l1'], cfg['cod']), (cfg['l2'], cfg['ver']), (titulo, cfg['vig'])]
        for i, (t, d) in enumerate(filas):
            yy = y - (i + 1) * h
            c.rect(x0 + w_logo, yy, w_mid, h)
            c.rect(x0 + w_logo + w_mid, yy, w_der, h)
            c.setFont('Helvetica-Bold', 7.6 if i < 2 else 8.6)
            c.drawCentredString(x0 + w_logo + w_mid / 2, yy + 5, t[:150])
            c.setFont('Helvetica-Bold', 7.6)
            c.drawCentredString(x0 + w_logo + w_mid + w_der / 2, yy + 5, d)
        yy = y - 3 * h - 14
        c.setFont('Helvetica-Bold', 8)
        partes = [('UNIDAD: ', serv), ('MES: ', nombre_mes(mid) if mid and '-' in mid and len(mid) == 7 else mid), ('BODEGA: ', 'EMERGENCIAS' if serv == 'EMERGENCIA' else serv), ('RESPONSABLE: ', cfg.get('resp') or '')]
        xx = x0
        paso = (ancho - 2 * margen) / 4
        for k, v in partes:
            c.setFont('Helvetica-Bold', 8)
            c.drawString(xx, yy, k)
            c.setFont('Helvetica', 8)
            c.drawString(xx + c.stringWidth(k, 'Helvetica-Bold', 8), yy, str(v))
            xx += paso
        return yy - 8

    def _doc(self, ruta, pagesize, titulo, serv, mid, top=118):
        R = _rl()
        ancho, alto = pagesize
        doc = R['SimpleDocTemplate'](ruta, pagesize=pagesize, leftMargin=26, rightMargin=26, topMargin=top, bottomMargin=34, title=titulo)
        pdf = self

        def pagina(c, d):
            c.saveState()
            pdf.encabezado(c, ancho, alto, titulo, serv, mid)
            c.restoreState()
        return doc, pagina

    def _estilo_tabla(self, fs=7, cab_rows=1, extra=None):
        R = _rl()
        col = R['colors']
        st = [('FONT', (0, 0), (-1, -1), 'Helvetica', fs), ('GRID', (0, 0), (-1, -1), 0.4, col.HexColor('#7f8a9c')),
              ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
              ('BACKGROUND', (0, 0), (-1, cab_rows - 1), col.HexColor('#41608f')), ('TEXTCOLOR', (0, 0), (-1, cab_rows - 1), col.white),
              ('FONT', (0, 0), (-1, cab_rows - 1), 'Helvetica-Bold', fs), ('TOPPADDING', (0, 0), (-1, -1), 1.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
              ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2)]
        return R['TableStyle'](st + (extra or []))

    def _p(self, txt, fs=7, bold=False, align=0):
        R = _rl()
        st = R['ParagraphStyle']('x', fontName='Helvetica-Bold' if bold else 'Helvetica', fontSize=fs, leading=fs + 1.6, alignment=align)
        return R['Paragraph'](str(txt).replace('&', '&amp;').replace('<', '&lt;'), st)

    def firmas(self):
        R = _rl()
        cfg = self.D.cfg
        if not cfg.get('firmas'):
            return R['Spacer'](1, 1)
        nombres = [x for x in (cfg.get('f1'), cfg.get('f2'), cfg.get('f3')) if x]
        if not nombres:
            return R['Spacer'](1, 1)
        w = 540 / len(nombres)
        t = R['Table']([[''] * len(nombres), nombres], colWidths=[w] * len(nombres), rowHeights=[34, 14])
        t.setStyle(R['TableStyle']([('LINEABOVE', (0, 1), (-1, 1), 0.8, R['colors'].black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONT', (0, 0), (-1, -1), 'Helvetica', 8)]))
        return t
        
        t = R['Table']([['', '', ''], [cfg['f1'], cfg['f2'], cfg['f3']]], colWidths=[180, 180, 180], rowHeights=[34, 14])
        t.setStyle(R['TableStyle']([('LINEABOVE', (0, 1), (0, 1), 0.8, R['colors'].black), ('LINEABOVE', (1, 1), (1, 1), 0.8, R['colors'].black),
                                    ('LINEABOVE', (2, 1), (2, 1), 0.8, R['colors'].black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONT', (0, 0), (-1, -1), 'Helvetica', 8)]))
        return t

    # ---- Matriz M03 del mes (oficio horizontal)
    def matriz(self, serv, mid):
        R = _rl()
        m = self.D.mes(serv, mid)
        c = self.D.calc_mes(serv, mid)
        ps = R['landscape'](R['legal'])
        ruta = os.path.join(self.carpeta(serv, mid), f'1. Matriz de entradas y salidas - {nombre_mes(mid)} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, self.D.cfg['tit'], serv, mid)
        cab = ['CÓDIGO', 'INSUMOS', 'EXIST.\nINICIAL', 'CANT.\nRECIB.', 'FECHA\nINGRESO', 'FECHA\nVENC.'] + [str(d) for d in range(1, 32)] + ['TOTAL\nCONS.', 'SALDO\nFINAL']
        filas = [cab]
        extra = []
        tc = tr = 0
        for it in m['items']:
            if it.get('t') == 'c':
                filas.append(['', it['n']] + [''] * 37)
                extra += [('SPAN', (1, len(filas) - 1), (-1, len(filas) - 1)), ('BACKGROUND', (0, len(filas) - 1), (-1, len(filas) - 1), R['colors'].HexColor('#d9e1f2')),
                          ('FONT', (0, len(filas) - 1), (-1, len(filas) - 1), 'Helvetica-Bold', 6.5)]
                continue
            x = c[it['id']]
            tc += x['cons']
            tr += x['rec']
            d = it.get('d') or []
            fila = [it.get('c', ''), self._p(it['n'], 5.6), fmt(x['ini']), fmt(x['rec']) if x['rec'] else '', ', '.join(f[8:10] + '/' + f[5:7] for f in x['fechas']), vto_txt(x['vto'])]
            fila += [fmt(d[k]) if k < len(d) and d[k] not in (None, '') else '' for k in range(31)]
            fila += [fmt(x['cons']), fmt(x['sal'])]
            filas.append(fila)
            if x['sal'] < 0:
                extra.append(('TEXTCOLOR', (-1, len(filas) - 1), (-1, len(filas) - 1), R['colors'].red))
        filas.append(['', 'TOTALES', '', fmt(tr), '', ''] + [''] * 31 + [fmt(tc), ''])
        extra += [('FONT', (0, len(filas) - 1), (-1, len(filas) - 1), 'Helvetica-Bold', 6.5), ('FONT', (-1, 1), (-1, -1), 'Helvetica-Bold', 6)]
        aw = ps[0] - 52
        fijos = [40, 0, 26, 26, 30, 28] + [16.2] * 31 + [26, 28]
        fijos[1] = aw - sum(fijos)
        t = R['Table'](filas, colWidths=fijos, repeatRows=1)
        t.setStyle(self._estilo_tabla(5.8, 1, extra))
        doc.build([t, R['Spacer'](1, 18), self.firmas()], onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Existencias (de un mes a otro)
    def existencias(self, serv, mid):
        R = _rl()
        m = self.D.mes(serv, mid)
        c = self.D.calc_mes(serv, mid)
        ps = R['landscape'](R['letter'])
        ruta = os.path.join(self.carpeta(serv, mid), f'Existencias - {nombre_mes(mid)} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, 'EXISTENCIA DE INSUMOS DE UN MES A OTRO', serv, mid)
        filas = [['CÓDIGO', 'INSUMO', 'SALDO MES\nANTERIOR', 'EXISTENCIA\nINICIAL', 'INGRESOS', 'CONSUMO', 'SALDO AL\nCIERRE', 'ESTADO']]
        extra = []
        for it in m['items']:
            if it.get('t') != 'i':
                continue
            x = c[it['id']]
            if x['sinDatos']:
                continue
            est = 'Negativo' if x['sal'] < 0 else ('Falta existencia inicial' if x['incompleto'] else ('Agotado' if x['sal'] == 0 else 'Con existencia'))
            ini = fmt(x['ini']) + (' (conteo)' if x['iniManual'] else '')
            filas.append([it.get('c', ''), self._p(it['n'], 6.5), '' if x['ant'] is None else fmt(x['ant']), ini, fmt(x['rec']) if x['rec'] else '', fmt(x['cons']) if x['cons'] else '', fmt(x['sal']), est])
            if x['sal'] < 0:
                extra.append(('TEXTCOLOR', (6, len(filas) - 1), (7, len(filas) - 1), R['colors'].red))
        aw = ps[0] - 52
        w = [52, 0, 58, 70, 52, 52, 56, 100]
        w[1] = aw - sum(w)
        t = R['Table'](filas, colWidths=w, repeatRows=1)
        t.setStyle(self._estilo_tabla(7, 1, extra))
        nota = self._p('Existencia inicial = saldo al cierre del mes anterior (automática); "(conteo)" = escrita a mano por conteo físico. Saldo = inicial + ingresos − consumo. Un saldo negativo no pasa al mes siguiente.', 7)
        doc.build([t, R['Spacer'](1, 8), nota, R['Spacer'](1, 18), self.firmas()], onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Matriz por período
    def periodo(self, serv, desde, hasta):
        R = _rl()
        filas_d, meses = self.app.consolidado(serv, desde, hasta)
        ps = R['landscape'](R['letter'])
        etiqueta = nombre_mes(desde) + (' a ' + nombre_mes(hasta) if hasta != desde else '')
        ruta = os.path.join(self.D.base, 'matrices', f'Matriz {desde} a {hasta} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, 'MATRIZ DE INSUMOS POR PERÍODO: ' + etiqueta, serv, etiqueta)
        cab = ['CÓDIGO', 'INSUMO', 'EXIST.\nINICIAL'] + [f'{nombre_mes(x)[:3]}\n{x[:4]}' for x in meses] + ['INGRESOS', 'CONSUMO\nTOTAL', 'SALDO\nFINAL']
        filas = [cab]
        for f in filas_d:
            filas.append([f['c'], self._p(f['n'], 6.2), fmt(f['ini']) if f['ini'] is not None else '—'] + [fmt(f['cons_mes'].get(x, '')) for x in meses] + [fmt(f['rec']), fmt(f['cons']), fmt(f['sal'])])
        aw = ps[0] - 52
        w = [50, 0, 44] + [38] * len(meses) + [46, 48, 46]
        w[1] = max(120, aw - sum(w))
        t = R['Table'](filas, colWidths=w, repeatRows=1)
        t.setStyle(self._estilo_tabla(6.6, 1))
        doc.build([t, R['Spacer'](1, 18), self.firmas()], onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Recibo de recepción
    def recibo(self, serv, mid, vale):
        R = _rl()
        m = self.D.mes(serv, mid)
        c = self.D.calc_mes(serv, mid)
        ps = R['letter']
        ruta = os.path.join(self.carpeta(serv, mid), f'Recibo {vale.get("rec", "").replace("/", "-")} vale {str(vale.get("num", "")).replace("/", "-")}.pdf')
        doc, pagina = self._doc(ruta, ps, 'RECIBO DE RECEPCIÓN DE INSUMOS', serv, mid)
        por = {it['id']: it for it in m['items']}
        info = R['Table']([[f'RECIBO N° {vale.get("rec") or "—"}', f'N° de vale: {vale.get("num", "")}', f'Fecha de recepción: {fecha_txt(vale.get("fecha", ""))}', f'Procedencia: {vale.get("proc") or "Almacén General"}']], colWidths=[130, 130, 160, 140])
        info.setStyle(R['TableStyle']([('GRID', (0, 0), (-1, -1), 0.5, R['colors'].grey), ('FONT', (0, 0), (-1, -1), 'Helvetica-Bold', 8)]))
        filas = [['#', 'CÓDIGO', 'INSUMO', 'LOTE / OBS.', 'VENCE', 'CANTIDAD\nRECIBIDA', 'SALDO DEL MES\n(con este ingreso)']]
        tot = 0
        for k, l in enumerate(vale['lineas'], 1):
            it = por.get(l['item'], {'n': '?'})
            tot += l.get('cant') or 0
            filas.append([k, it.get('c', ''), self._p(it.get('n', ''), 7), l.get('obs', ''), vto_txt(l.get('vto', '')), fmt(l.get('cant')), fmt(c.get(l['item'], {}).get('sal'))])
        filas.append(['', '', 'TOTAL', '', '', fmt(tot), ''])
        t = R['Table'](filas, colWidths=[18, 52, 250, 80, 42, 52, 66], repeatRows=1)
        t.setStyle(self._estilo_tabla(7.5, 1, [('FONT', (0, len(filas) - 1), (-1, len(filas) - 1), 'Helvetica-Bold', 7.5)]))
        txt = self._p('Se hace constar que se recibieron en la bodega los insumos detallados, revisando cantidad, lote y fecha de vencimiento.', 8)
        f = R['Table']([['', ''], ['Entregó (Almacén General)', 'Recibió (Bodega)']], colWidths=[270, 270], rowHeights=[40, 14])
        f.setStyle(R['TableStyle']([('LINEABOVE', (0, 1), (-1, 1), 0.8, R['colors'].black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONT', (0, 0), (-1, -1), 'Helvetica', 8)]))
        if not self.D.cfg.get('firmas'):
            f = R['Spacer'](1, 1)
        doc.build([info, R['Spacer'](1, 8), t, R['Spacer'](1, 10), txt, R['Spacer'](1, 10), f], onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Hojas de pedido / retiro (una por día, carta vertical, letra grande)
    def hojas(self, serv, mid, dias, disp=True):
        R = _rl()
        cv = R['rlcanvas']
        m = self.D.mes(serv, mid)
        c = self.D.calc_mes(serv, mid)
        its = [it for it in m['items'] if it.get('t') == 'i']
        nombre = f'Hojas de pedido {dias[0]:02d}' + (f' al {dias[-1]:02d}' if len(dias) > 1 else '') + f' {nombre_mes(mid)} {serv}.pdf'
        ruta = os.path.join(self.carpeta(serv, mid), nombre)
        W, H = R['letter']
        pdf = cv.Canvas(ruta, pagesize=(W, H))
        M = 30
        FULL, ULT, RH = 38, 28, 16.2
        lista = its + [None] * 4
        pags, resto = [], list(lista)
        while len(resto) > ULT * 2:
            pags.append(resto[:min(FULL * 2, len(resto))])
            resto = resto[min(FULL * 2, len(resto)):]
        pags.append(resto)
        colw = (W - 2 * M - 14) / 2
        CW, QW, DW = 46, 62, (40 if disp else 0)
        NW = colw - CW - QW - DW
        for dia in dias:
            fecha = f'{mid}-{dia:02d}'
            folio = self.D.folio_hoja(serv, fecha)
            ds = DSEM[datetime.date(*map(int, fecha.split('-'))).weekday()]
            for n, pg in enumerate(pags):
                ult = n == len(pags) - 1
                filas = (len(pg) + 1) // 2 if ult else FULL
                pdf.setFont('Times-Bold', 13)
                hosp = self.D.cfg.get('hospital', '').split(',')
                pdf.drawString(M, H - M - 8, hosp[0].strip() + ',')
                pdf.drawString(M, H - M - 24, (hosp[1].strip() if len(hosp) > 1 else '') + f',    SERVICIO: {serv}')
                pdf.setFont('Helvetica', 10)
                pdf.drawRightString(W - M, H - M - 6, 'HOJA N°')
                pdf.setFont('Helvetica-Bold', 15)
                pdf.drawRightString(W - M, H - M - 22, f'{folio:05d}')
                pdf.setFont('Helvetica', 8.5)
                pdf.drawRightString(W - M, H - M - 34, f'FECHA: {fecha_txt(fecha)} ({ds}) · pág. {n + 1} de {len(pags)}')
                y0 = H - M - 52
                for col in range(2):
                    x = M + col * (colw + 14)
                    nfil = filas if col == 0 else (max(len(pg) - filas, filas if ult else FULL))
                    # encabezado de columnas
                    pdf.setFillGray(0.88)
                    pdf.rect(x, y0 - 22, colw, 22, fill=1, stroke=0)
                    pdf.setFillGray(0)
                    pdf.setLineWidth(1)
                    pdf.setFont('Helvetica-Bold', 7.5)
                    xs = [x, x + CW, x + CW + NW, x + CW + NW + DW, x + colw]
                    tits = ['CÓDIGO', 'INSUMO'] + (['DISPON.'] if disp else []) + ['CANTIDAD RETIRADA']
                    bordes = [xs[0], xs[1], xs[2]] + ([xs[3]] if disp else []) + [xs[4]]
                    for k, ttxt in enumerate(tits):
                        cx = (bordes[k] + bordes[k + 1]) / 2
                        pdf.drawCentredString(cx, y0 - 14, ttxt)
                    for r in range(nfil + 1):
                        yy = y0 - 22 - r * RH
                        pdf.line(x, yy, x + colw, yy)
                    pdf.line(x, y0, x + colw, y0)
                    for bx in bordes:
                        pdf.line(bx, y0, bx, y0 - 22 - nfil * RH)
                    for r in range(nfil):
                        j = r if col == 0 else filas + r
                        it = pg[j] if j < len(pg) else None
                        if col == 0 and r >= filas:
                            it = None
                        if not it:
                            continue
                        yy = y0 - 22 - (r + 1) * RH + 5
                        pdf.setFont('Helvetica', 7.8)
                        pdf.drawCentredString((xs[0] + xs[1]) / 2, yy, it.get('c', ''))
                        pdf.setFont('Times-Roman', 10.5)
                        txt = nombre_hoja(it)
                        while pdf.stringWidth(txt, 'Times-Roman', 10.5) > NW - 6 and len(txt) > 4:
                            txt = txt[:-2] + '…' if not txt.endswith('…') else txt[:-2] + '…'
                        pdf.drawString(xs[1] + 3, yy, txt)
                        if disp:
                            x_ = c[it['id']]
                            if not (x_['sinDatos'] or x_['incompleto']):
                                rec = sum(l['cant'] for v in m['vales'] if v.get('fecha', '') <= fecha for l in v['lineas'] if l['item'] == it['id'])
                                cons = sum(v or 0 for v in (it.get('d') or [])[:dia - 1] if isinstance(v, (int, float)))
                                dd = (x_['ini'] or 0) + rec - cons
                                pdf.setFont('Helvetica', 7.5)
                                if dd == 0:
                                    pdf.setFillColorRGB(0.75, 0, 0)
                                    pdf.drawCentredString((xs[2] + xs[3]) / 2, yy, 'AGOTADO')
                                    pdf.setFillGray(0)
                                elif dd > 0:
                                    pdf.drawCentredString((xs[2] + xs[3]) / 2, yy, fmt(dd))
                if ult:
                    yb = y0 - 22 - filas * RH - 28
                    pdf.setFont('Times-Bold', 11)
                    for k, (lab, val) in enumerate([('SERVICIO:', serv), ('FECHA:', fecha_txt(fecha)), ('RESPONSABLE (nombre de quien retira):', ''), ('FIRMA:', '')]):
                        yy = yb - k * 30
                        pdf.setFont('Times-Bold', 11)
                        pdf.drawString(M, yy, lab)
                        lx = M + pdf.stringWidth(lab, 'Times-Bold', 11) + 8
                        pdf.line(lx, yy - 2, M + 360, yy - 2)
                        pdf.setFont('Helvetica', 10)
                        pdf.drawString(lx + 4, yy + 1, val)
                    pdf.setDash(3, 2)
                    pdf.rect(M + 380, yb - 100, W - 2 * M - 380, 112)
                    pdf.setDash()
                    pdf.setFont('Helvetica', 8)
                    pdf.drawCentredString(M + 380 + (W - 2 * M - 380) / 2, yb - 94, 'SELLO')
                    pdf.setFont('Helvetica', 7)
                    pdf.drawString(M, yb - 118, ('"Dispon." = existencia al inicio del día según el sistema. ' if disp else '') + 'Escriba en CANTIDAD RETIRADA el número que retira. Al registrar esta hoja se descuenta de la existencia.')
                pdf.setFont('Helvetica', 7)
                pdf.drawCentredString(W / 2, 16, f'Hoja N° {folio:05d} · {fecha_txt(fecha)} · página {n + 1} de {len(pags)}')
                pdf.showPage()
        pdf.save()
        return ruta

    # ---- Control de retiros del mes
    def retiros(self, serv, mid):
        R = _rl()
        m = self.D.mes(serv, mid)
        por = {it['id']: it for it in m['items']}
        ps = R['letter']
        ruta = os.path.join(self.carpeta(serv, mid), f'Control de retiros - {nombre_mes(mid)} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, 'CONTROL DE RETIROS DE BODEGA', serv, mid)
        filas = [['N° HOJA', 'FECHA', 'RESPONSABLE', 'INSUMOS RETIRADOS', 'UNID.']]
        tot = {}
        for r in sorted(m['retiros'], key=lambda r: (r.get('fecha', ''), str(r.get('folio', '')))):
            u = sum(l['cant'] for l in r['lineas'])
            txt = ' · '.join(f"{nombre_hoja(por.get(l['item'], {'n': '?'}))} {fmt(l['cant'])}" for l in r['lineas'])
            filas.append([r.get('folio', ''), fecha_txt(r.get('fecha', '')), self._p(r.get('resp', ''), 7), self._p(txt, 6.5), fmt(u) + ('' if r.get('desc', True) else '\n(sin desc.)')])
            for l in r['lineas']:
                tot[l['item']] = tot.get(l['item'], 0) + l['cant']
        if len(filas) == 1:
            filas.append(['', '', 'No hay retiros registrados', '', ''])
        t = R['Table'](filas, colWidths=[44, 52, 110, 314, 40], repeatRows=1)
        t.setStyle(self._estilo_tabla(7.2))
        f2 = [['CÓDIGO', 'INSUMO', 'TOTAL RETIRADO']] + [[por[i].get('c', ''), self._p(por[i]['n'], 7), fmt(v)] for i, v in sorted(tot.items(), key=lambda kv: -kv[1]) if i in por]
        t2 = R['Table'](f2 if len(f2) > 1 else f2 + [['', 'Sin datos', '']], colWidths=[60, 420, 80], repeatRows=1)
        t2.setStyle(self._estilo_tabla(7.2))
        doc.build([t, R['Spacer'](1, 12), self._p('Total retirado por insumo', 9, True), R['Spacer'](1, 4), t2, R['Spacer'](1, 18), self.firmas()],
                  onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Tarjetas de control de existencia (kardex)
    def kardex(self, serv, mid):
        R = _rl()
        ps = R['landscape'](R['letter'])
        ruta = os.path.join(self.carpeta(serv, mid), f'Tarjetas de control (kardex) hasta {nombre_mes(mid)} {serv}.pdf')
        doc = R['SimpleDocTemplate'](ruta, pagesize=ps, leftMargin=26, rightMargin=26, topMargin=26, bottomMargin=30, title='Tarjetas de control de existencia')
        historia = self.app.historia(serv, mid)
        story = []
        n = 0
        m = self.D.mes(serv, mid)
        for it in m['items']:
            if it.get('t') != 'i':
                continue
            movs = historia.get(clave(it), [])
            if not any(x.get('ent') or x.get('sal') for x in movs):
                continue
            n += 1
            cab = R['Table']([[f'CÓDIGO ARTÍCULO: {it.get("c", "")}', 'MINISTERIO DE SALUD — CONTROL DE EXISTENCIA DE MEDICAMENTOS E INSUMOS MÉDICOS', f'TARJETA N° {n}'],
                              [self._p('NOMBRE ARTÍCULO: ' + it['n'], 8), 'ALMACÉN: Bodega de ' + serv.title(), '']], colWidths=[250, 380, 110])
            cab.setStyle(R['TableStyle']([('FONT', (0, 0), (-1, -1), 'Helvetica-Bold', 8), ('BOX', (1, 0), (1, 1), 1, R['colors'].black), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
            filas = [['FECHA', 'CONCEPTO', 'PROCEDENCIA', 'N° LOTE', 'VENCE', 'ENTRADA', 'SALIDA', 'SALDO']]
            for x in movs:
                filas.append([fecha_txt(x['f']), self._p(x['con'], 7), x.get('proc', ''), x.get('lote', ''), vto_txt(x.get('vto', '')), fmt(x.get('ent')) if x.get('ent') else '', fmt(x.get('sal')) if x.get('sal') else '', fmt(x['saldo'])])
            while len(filas) < 14:
                filas.append([''] * 8)
            ent = sum(x.get('ent') or 0 for x in movs)
            sal = sum(x.get('sal') or 0 for x in movs)
            filas.append(['', 'TOTALES', '', '', '', fmt(ent), fmt(sal), fmt(movs[-1]['saldo'] if movs else 0)])
            t = R['Table'](filas, colWidths=[60, 270, 90, 70, 50, 65, 65, 70], repeatRows=1, rowHeights=[20] + [22] * (len(filas) - 1))
            t.setStyle(self._estilo_tabla(8, 1, [('FONT', (0, len(filas) - 1), (-1, len(filas) - 1), 'Helvetica-Bold', 8)]))
            story += [cab, R['Spacer'](1, 6), t, R['PageBreak']()]
        if not story:
            story = [self._p('No hay insumos con movimiento.', 10)]
        doc.build(story, canvasmaker=self._num_canvas())
        return ruta

    # ---- Informe para auditoría interna y externa
    def auditoria(self, serv, desde, hasta):
        R = _rl()
        meses = [x for x in self.D.meses(serv) if desde <= x <= hasta]
        etiqueta = nombre_mes(desde) + (' a ' + nombre_mes(hasta) if hasta != desde else '')
        ps = R['letter']
        ruta = os.path.join(self.D.base, 'documentos', f'Informe para auditoria {desde} a {hasta} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, 'INFORME DE MOVIMIENTO DE INSUMOS PARA AUDITORÍA INTERNA Y EXTERNA', serv, etiqueta)
        todo = self.D.calcular(serv, hasta)
        P = lambda t, fs=8.5, b=False: self._p(t, fs, b)
        story = [P('1. OBJETIVO Y ALCANCE', 10, True),
                 P(f'Presentar a la auditoría interna y externa los movimientos de insumos de la bodega de {serv.title()} registrados en la matriz {self.D.cfg["cod"]} durante {etiqueta}: existencia inicial, ingresos del Almacén General, consumo y saldo final, con sus documentos de respaldo y las observaciones encontradas.'),
                 R['Spacer'](1, 6), P('2. FUENTES Y MÉTODO', 10, True),
                 P('Matriz de entradas y salidas de insumos (por mes y día); vales de salida del Almacén General y recibos; hojas diarias de pedido/retiro; conteos físicos. Saldo final = existencia inicial + ingresos − consumo. La existencia inicial de cada mes es el saldo del mes anterior, salvo conteo físico.'),
                 R['Spacer'](1, 6), P('3. RESUMEN POR MES', 10, True)]
        filas = [['MES', 'INSUMOS', 'VALES', 'UNID. RECIBIDAS', 'UNID. CONSUMIDAS', 'HOJAS RETIRO', 'CONTADOS', 'SALDOS NEG.']]
        nvales = 0
        for mid in meses:
            m = self.D.mes(serv, mid)
            c = todo[mid]
            its = [i for i in m['items'] if i.get('t') == 'i']
            nvales += len(m['vales'])
            filas.append([nombre_mes(mid), len(its), len(m['vales']), fmt(sum(c[i['id']]['rec'] for i in its)), fmt(sum(c[i['id']]['cons'] for i in its)), len(m['retiros']),
                          len([v for v in m['conteo'].values() if v is not None]), sum(1 for i in its if c[i['id']]['sal'] < 0)])
        t = R['Table'](filas, repeatRows=1)
        t.setStyle(self._estilo_tabla(7.5))
        story += [t, R['Spacer'](1, 8), P('4. MOVIMIENTO CONSOLIDADO POR INSUMO', 10, True)]
        cons, _ = self.app.consolidado(serv, desde, hasta)
        f2 = [['CÓDIGO', 'INSUMO', 'EXIST. INICIAL', 'INGRESOS', 'CONSUMO', 'SALDO FINAL', 'OBSERVACIÓN']]
        sin_ini = neg = sin_cod = 0
        for o in cons:
            ob = []
            if not o['c']:
                ob.append('Sin código')
                sin_cod += 1
            if o['ini'] is None and o['cons'] > 0:
                ob.append('Sin existencia inicial')
                sin_ini += 1
            if o['sal'] < 0:
                ob.append('Saldo negativo')
                neg += 1
            f2.append([o['c'], self._p(o['n'], 6.5), '—' if o['ini'] is None else fmt(o['ini']), fmt(o['rec']), fmt(o['cons']), fmt(o['sal']), self._p('; '.join(ob), 6.5)])
        t2 = R['Table'](f2, colWidths=[48, 240, 46, 44, 44, 46, 92], repeatRows=1)
        t2.setStyle(self._estilo_tabla(6.8))
        story += [t2, R['Spacer'](1, 8), P('5. INGRESOS Y SU RESPALDO', 10, True)]
        f3 = [['FECHA', 'N° VALE', 'N° RECIBO', 'CÓDIGO', 'INSUMO', 'CANT.', 'VENCE']]
        for mid in meses:
            m = self.D.mes(serv, mid)
            por = {i['id']: i for i in m['items']}
            for v in sorted(m['vales'], key=lambda v: v.get('fecha', '')):
                for l in v['lineas']:
                    it = por.get(l['item'], {'n': '?'})
                    f3.append([fecha_txt(v.get('fecha', '')), v.get('num', ''), v.get('rec', '') or '—', it.get('c', ''), self._p(it.get('n', ''), 6.5), fmt(l.get('cant')), vto_txt(l.get('vto', ''))])
        t3 = R['Table'](f3 if len(f3) > 1 else f3 + [['', '', '', '', 'Sin ingresos registrados', '', '']], colWidths=[48, 52, 50, 48, 262, 40, 40], repeatRows=1)
        t3.setStyle(self._estilo_tabla(6.8))
        story += [t3, R['Spacer'](1, 8), P('6. OBSERVACIONES PARA LA AUDITORÍA', 10, True),
                  P(f'• {sin_ini} insumo(s) con consumo sin existencia inicial registrada: su saldo no es confiable hasta registrar un conteo físico.'),
                  P(f'• {neg} insumo(s) con saldo negativo al cierre: falta existencia inicial o vales de ingreso.'),
                  P(f'• {sin_cod} insumo(s) sin código institucional.'),
                  P(f'• Vales registrados en el período: {nvales}.'),
                  P('Observaciones del responsable: ______________________________________________________________________________'),
                  P('______________________________________________________________________________________________________________'),
                  R['Spacer'](1, 6), P('7. DECLARACIÓN', 10, True),
                  P('Declaro que la información de este informe fue generada a partir de los registros de la bodega del período indicado y que los documentos de respaldo se ponen a disposición de la auditoría.'),
                  R['Spacer'](1, 26)]
        f = R['Table']([['', '', ''], ['Elaboró (Responsable de bodega)', 'Revisó (Jefatura)', 'Recibido por auditoría (nombre, firma, fecha)']], colWidths=[180, 180, 200], rowHeights=[30, 14])
        f.setStyle(R['TableStyle']([('LINEABOVE', (0, 1), (-1, 1), 0.8, R['colors'].black), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONT', (0, 0), (-1, -1), 'Helvetica', 7.5)]))
        story.append(f)
        doc.build(story, onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta

    # ---- Informe mensual de control antisoborno
    def antisoborno(self, serv, mid):
        R = _rl()
        m = self.D.mes(serv, mid)
        c = self.D.calc_mes(serv, mid)
        por = {i['id']: i for i in m['items']}
        ps = R['letter']
        ruta = os.path.join(self.carpeta(serv, mid), f'Informe antisoborno - {nombre_mes(mid)} {serv}.pdf')
        doc, pagina = self._doc(ruta, ps, 'INFORME MENSUAL DE CONTROL ANTISOBORNO', serv, mid)
        P = lambda t, fs=8.5, b=False: self._p(t, fs, b)
        sinvale = [(v, l) for v in m['vales'] for l in v['lineas'] if not v.get('num') or re.search(r'(?i)directo|s/?n', v.get('num', ''))]
        vence = []
        for v in m['vales']:
            for l in v['lineas']:
                if not l.get('vto'):
                    vence.append((v, l, 'Sin fecha de vencimiento'))
                elif v.get('fecha'):
                    a, b = map(int, v['fecha'][:7].split('-'))
                    cc, dd = map(int, l['vto'][:7].split('-'))
                    if (cc - a) * 12 + (dd - b) < 6:
                        vence.append((v, l, 'Menos de 6 meses de vida útil'))
        neg = [i for i in m['items'] if i.get('t') == 'i' and c[i['id']]['sal'] < 0]
        aj = [i for i in m['items'] if i.get('t') == 'i' and c[i['id']]['ajuste'] is not None]
        dif = [(i, m['conteo'][i['id']], c[i['id']]['sal']) for i in m['items'] if i.get('t') == 'i' and m['conteo'].get(i['id']) is not None and m['conteo'][i['id']] != c[i['id']]['sal']]
        atip = []
        for i in m['items']:
            if i.get('t') != 'i':
                continue
            ds = [(k + 1, v) for k, v in enumerate(i.get('d') or []) if isinstance(v, (int, float)) and v > 0]
            if len(ds) >= 3:
                med = sorted(v for _, v in ds)[len(ds) // 2]
                atip += [(i, d, v, med) for d, v in ds if v >= 3 * med and v - med >= 20]
        cats = [('Ingresos sin número de vale', len(sinvale)), ('Recibidos con vencimiento corto o sin vencimiento', len(vence)), ('Diferencias en conteo físico', len(dif)),
                ('Ajustes de existencia inicial', len(aj)), ('Saldos negativos', len(neg)), ('Consumos fuera de lo normal', len(atip))]
        t = R['Table']([['CONTROL', 'SEÑALES', 'RESULTADO']] + [[a, b, 'Revisar' if b else 'Conforme'] for a, b in cats], colWidths=[330, 70, 80])
        t.setStyle(self._estilo_tabla(8))
        det = []
        for v, l, mot in vence:
            det.append([v.get('num', ''), self._p(por.get(l['item'], {}).get('n', ''), 6.8), vto_txt(l.get('vto', '')), mot])
        for i in neg:
            det.append(['', self._p(i['n'], 6.8), '', f"Saldo negativo {fmt(c[i['id']]['sal'])}"])
        for i, d, v, med in atip:
            det.append(['', self._p(i['n'], 6.8), '', f'Día {d}: {fmt(v)} (usual {fmt(med)})'])
        t2 = R['Table']([['VALE', 'INSUMO', 'VENCE', 'DETALLE']] + (det or [['', 'Sin señales', '', '']]), colWidths=[50, 300, 45, 165], repeatRows=1)
        t2.setStyle(self._estilo_tabla(7))
        a = m.get('antisoborno') or {}
        story = [P('1. OBJETIVO', 10, True), P(f'Revisar que las entradas, salidas y existencias de {nombre_mes(mid)} estén respaldadas y dejar constancia de las señales de alerta y su revisión.'),
                 R['Spacer'](1, 6), P('2. CONTROLES', 10, True), t, R['Spacer'](1, 6), P('3. DETALLE DE SEÑALES DE ALERTA', 10, True), t2, R['Spacer'](1, 6),
                 P('4. RESULTADO DE LA REVISIÓN', 10, True), P(a.get('concl') or '(sin completar)'), P('5. OBSERVACIONES Y JUSTIFICACIÓN', 10, True), P(a.get('obs') or ''),
                 P('6. ACCIONES', 10, True), P(a.get('acc') or ''), R['Spacer'](1, 6), P('7. DECLARACIÓN', 10, True),
                 P(f'Declaro que la información corresponde a los registros de la bodega de {nombre_mes(mid)}; que los ingresos están respaldados por los vales indicados; que en el manejo de los insumos no he ofrecido, solicitado ni aceptado dinero, regalos, favores ni otro beneficio; y que las señales de alerta fueron revisadas.'),
                 P('Las señales de alerta son controles automáticos; por sí solas no significan una falta.', 7), R['Spacer'](1, 20), self.firmas()]
        doc.build(story, onFirstPage=pagina, onLaterPages=pagina, canvasmaker=self._num_canvas())
        return ruta


# --------------------------------------------------------------------------- ventana principal
class App(tk.Tk):
    DIAS = [f'd{k}' for k in range(1, 32)]

    def __init__(self, base=None):
        super().__init__()
        self.D = Datos(base)
        self.pdf = PDF(self)
        self.title(f'{APP} — versión {VERSION}')
        self.geometry('1360x780')
        try:
            self.iconbitmap(recurso('bodega.ico'))
        except Exception:
            pass
        st = ttk.Style(self)
        try:
            st.theme_use('clam')
        except Exception:
            pass
        st.configure('Treeview', rowheight=22, font=('Segoe UI', 9))
        st.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'), background='#41608f', foreground='white')
        st.map('Treeview.Heading', background=[('active', '#2c4670')])
        st.map('Treeview', background=[('selected', '#cfe0ff')], foreground=[('selected', 'black')])
        st.configure('TNotebook.Tab', padding=(12, 5), font=('Segoe UI', 9, 'bold'))
        self.serv = tk.StringVar()
        self.mid = tk.StringVar()
        self._barra()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=6, pady=4)
        self._tab_matriz()
        self._tab_ingresos()
        self._tab_retiros()
        self._tab_existencias()
        self._tab_periodo()
        self._tab_modelos()
        self._tab_documentos()
        self._tab_ajustes()
        self.estado = ttk.Label(self, text='', anchor='w')
        self.estado.pack(fill='x', padx=8, pady=(0, 4))
        self.nb.bind('<<NotebookTabChanged>>', lambda e: self.refrescar())
        self.after(100, self._inicio)

    # ---------------- barra superior
    def _barra(self):
        b = ttk.Frame(self)
        b.pack(fill='x', padx=6, pady=(6, 0))
        try:
            self._logo = tk.PhotoImage(file=recurso('logo.png')).subsample(3, 3)
            ttk.Label(b, image=self._logo).pack(side='left', padx=(0, 8))
        except Exception:
            pass
        ttk.Label(b, text='Sistema de Bodega de Emergencias', font=('Segoe UI', 13, 'bold'), foreground='#1f3a68').pack(side='left')
        ttk.Label(b, text=f'  v{VERSION}', foreground='#666').pack(side='left')
        ttk.Label(b, text='   Servicio:').pack(side='left')
        self.cb_serv = ttk.Combobox(b, textvariable=self.serv, width=16, state='readonly')
        self.cb_serv.pack(side='left')
        self.cb_serv.bind('<<ComboboxSelected>>', lambda e: self._cambio_serv())
        ttk.Button(b, text='+ Servicio', command=self.nuevo_servicio).pack(side='left', padx=2)
        ttk.Label(b, text='   Mes:').pack(side='left')
        self.cb_mes = ttk.Combobox(b, textvariable=self.mid, width=18, state='readonly')
        self.cb_mes.pack(side='left')
        self.cb_mes.bind('<<ComboboxSelected>>', lambda e: self.refrescar())
        ttk.Button(b, text='+ Mes siguiente', command=self.nuevo_mes).pack(side='left', padx=2)
        ttk.Button(b, text='📂 Carpeta de datos', command=lambda: abrir(self.D.base)).pack(side='right')

    def _inicio(self):
        servs = self.D.servicios()
        self.cb_serv['values'] = servs
        self.serv.set(servs[0])
        if not self.D.meses(self.serv.get()):
            if messagebox.askyesno(APP, 'Bienvenido. No hay datos todavía.\n\n¿Desea IMPORTAR el respaldo del programa anterior?\n'
                                        '(En el programa anterior: Respaldo y ajustes > Descargar respaldo)\n\n"No" = empezar con un mes vacío.'):
                self.importar()
            if not self.D.meses(self.serv.get()):
                hoy = datetime.date.today()
                self.D.crear_mes(self.serv.get(), f'{hoy.year}-{hoy.month:02d}', [])
        self._cambio_serv()

    def _cambio_serv(self):
        ms = self.D.meses(self.serv.get())
        self.cb_mes['values'] = [f'{nombre_mes(x)}  ({x})' for x in ms]
        hoy = datetime.date.today()
        actual = f'{hoy.year}-{hoy.month:02d}'
        sel = actual if actual in ms else (ms[-1] if ms else '')
        self.mid.set(f'{nombre_mes(sel)}  ({sel})' if sel else '')
        self.refrescar()

    def M(self):
        v = self.mid.get()
        m = re.search(r'\((\d{4}-\d{2})\)', v)
        return m.group(1) if m else None

    def mes_actual(self):
        return self.D.mes(self.serv.get(), self.M())

    def guardar(self, mid=None):
        self.D.guardar_mes(self.serv.get(), mid or self.M())
        self.estado.config(text=f'✔ Guardado {datetime.datetime.now():%H:%M:%S} en {self.D.ruta_mes(self.serv.get(), mid or self.M())}')

    def refrescar(self):
        if not self.M():
            return
        tab = self.nb.index(self.nb.select()) if self.nb.tabs() else 0
        [self.pintar_matriz, self.pintar_ingresos, self.pintar_retiros, self.pintar_existencias, self.pintar_periodo, self.pintar_modelo, lambda: None, self.pintar_ajustes][tab]()

    # ---------------- crear mes / servicio
    def nuevo_mes(self):
        serv = self.serv.get()
        ms = self.D.meses(serv)
        sig = mes_siguiente(ms[-1]) if ms else f'{datetime.date.today():%Y-%m}'
        mod = self.D.modelo(serv)['items']
        base = mod if mod else [{'t': i['t'], 'n': i['n'], 'c': i.get('c', '')} for i in self.D.mes(serv, ms[-1])['items']] if ms else []
        if not messagebox.askyesno(APP, f'Se creará {nombre_mes(sig)} para {serv} con {sum(1 for i in base if i["t"] == "i")} insumos '
                                        f'({"del modelo" if mod else "del mes anterior"}).\nLa existencia inicial se toma sola del saldo del mes anterior.\n\n¿Continuar?'):
            return
        items = []
        for i in base:
            o = {'id': nuevo_id(), 't': i['t'], 'n': i['n']}
            if i.get('c'):
                o['c'] = i['c']
            if i['t'] == 'i':
                o['d'] = []
            items.append(o)
        self.D.crear_mes(serv, sig, items)
        self._cambio_serv()
        self.mid.set(f'{nombre_mes(sig)}  ({sig})')
        self.refrescar()

    def nuevo_servicio(self):
        n = simpledialog.askstring(APP, 'Nombre del servicio nuevo (ej.: OBSERVACION, PEDIATRIA):', parent=self)
        if not n:
            return
        n = re.sub(r'[^A-Z0-9 ]', '', nr_txt(n)).strip().replace(' ', '_')
        if not n:
            return
        copiar = messagebox.askyesno(APP, f'¿Copiar la lista de insumos (modelo) de {self.serv.get()} para {n}?')
        mod = self.D.modelo(self.serv.get()) if copiar else {'items': []}
        self.D.guardar_modelo(n, {'servicio': n, 'items': mod['items']})
        mid = f'{datetime.date.today():%Y-%m}'
        self.D.crear_mes(n, mid, [dict({'id': nuevo_id(), 't': i['t'], 'n': i['n']}, **({'c': i['c']} if i.get('c') else {}), **({'d': []} if i['t'] == 'i' else {})) for i in mod['items']])
        self.cb_serv['values'] = self.D.servicios()
        self.serv.set(n)
        self._cambio_serv()

    # ================= PESTAÑA MATRIZ (consumo diario) =================
    def _tab_matriz(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='📅 Matriz / consumo diario')
        top = ttk.Frame(f)
        top.pack(fill='x', pady=4)
        ttk.Label(top, text='Clic en una casilla y escriba el número · Enter = baja · Tab = derecha · Supr = borrar · doble clic en código o insumo = corregir',
                  foreground='#555').pack(side='left')
        bar = ttk.Frame(f)
        bar.pack(fill='x', pady=(0, 4))
        ttk.Button(bar, text='➕ Agregar insumo', command=self.agregar_insumo).pack(side='left', padx=2)
        ttk.Button(bar, text='＋ Grupo', command=self.agregar_grupo).pack(side='left', padx=2)
        ttk.Button(bar, text='✏️ Editar insumo', command=self.editar_insumo).pack(side='left', padx=2)
        ttk.Separator(bar, orient='vertical').pack(side='left', fill='y', padx=6)
        ttk.Button(bar, text='▲ Subir', command=lambda: self.mover_insumo(-1)).pack(side='left', padx=2)
        ttk.Button(bar, text='▼ Bajar', command=lambda: self.mover_insumo(1)).pack(side='left', padx=2)
        ttk.Button(bar, text='↪ Mover a otro grupo', command=self.mover_a_grupo).pack(side='left', padx=2)
        ttk.Button(bar, text='A→Z Ordenar grupo por nombre', command=lambda: self.ordenar_grupo('n')).pack(side='left', padx=2)
        ttk.Button(bar, text='Ordenar grupo por código', command=lambda: self.ordenar_grupo('c')).pack(side='left', padx=2)
        self.v_orden_todos = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text='Aplicar el orden también a los meses siguientes y al modelo', variable=self.v_orden_todos).pack(side='left', padx=8)
        ttk.Button(bar, text='🗑 Quitar insumo', command=self.quitar_insumo).pack(side='right', padx=2)
        self.lb_pos = ttk.Label(f, text='', font=('Segoe UI', 10, 'bold'), foreground='#b45f00')
        self.lb_pos.pack(fill='x')
        cont = ttk.Frame(f)
        cont.pack(fill='both', expand=True)
        self.tl = ttk.Treeview(cont, columns=('cod', 'nom'), show='headings', selectmode='browse')
        self.tl.heading('cod', text='CÓDIGO')
        self.tl.heading('nom', text='INSUMO')
        self.tl.column('cod', width=74, anchor='center', stretch=False)
        self.tl.column('nom', width=330, stretch=False)
        cols = ['ini', 'rec', 'fing', 'vto'] + self.DIAS + ['tot', 'sal']
        self.tr = ttk.Treeview(cont, columns=cols, show='headings', selectmode='browse')
        heads = {'ini': 'EXIST.\nINICIAL', 'rec': 'RECIB.', 'fing': 'F. INGR.', 'vto': 'VENCE', 'tot': 'TOTAL', 'sal': 'SALDO'}
        for cname in cols:
            self.tr.heading(cname, text=heads.get(cname, cname[1:]))
            self.tr.column(cname, width={'ini': 62, 'rec': 56, 'fing': 64, 'vto': 58, 'tot': 58, 'sal': 62}.get(cname, 34), anchor='center', stretch=False)
        vs = ttk.Scrollbar(cont, orient='vertical', command=self._yview)
        hs = ttk.Scrollbar(cont, orient='horizontal', command=self.tr.xview)
        self.tl.configure(yscrollcommand=lambda a, b: self._ysync(vs, a, b))
        self.tr.configure(yscrollcommand=lambda a, b: self._ysync(vs, a, b), xscrollcommand=hs.set)
        self.tl.grid(row=0, column=0, sticky='ns')
        self.tr.grid(row=0, column=1, sticky='nsew')
        vs.grid(row=0, column=2, sticky='ns')
        hs.grid(row=1, column=1, sticky='ew')
        cont.columnconfigure(1, weight=1)
        cont.rowconfigure(0, weight=1)
        for t in (self.tl, self.tr):
            t.tag_configure('cat', background='#d9e1f2', font=('Segoe UI', 9, 'bold'))
            t.tag_configure('neg', foreground='#c00000')
            t.bind('<MouseWheel>', self._rueda)
            t.bind('<Button-4>', lambda e: self._rueda(e, -1))
            t.bind('<Button-5>', lambda e: self._rueda(e, 1))
        self.tl.bind('<<TreeviewSelect>>', lambda e: self._selsync(self.tl, self.tr))
        self.tr.bind('<<TreeviewSelect>>', lambda e: self._selsync(self.tr, self.tl))
        self.tr.bind('<Button-1>', self._clic_celda)
        self.tr.bind('<Double-1>', lambda e: self.editar_celda())
        self.tr.bind('<Return>', lambda e: self.editar_celda())
        self.tr.bind('<Key>', self._tecla)
        self.tl.bind('<Double-1>', self._editar_nombre)
        for t in (self.tl, self.tr):
            t.bind('<Alt-Up>', lambda e: (self.mover_insumo(-1), 'break')[1])
            t.bind('<Alt-Down>', lambda e: (self.mover_insumo(1), 'break')[1])
        self.cur_col = 'd1'
        self.editor = None

    def _yview(self, *a):
        self.tl.yview(*a)
        self.tr.yview(*a)

    def _ysync(self, sb, a, b):
        sb.set(a, b)
        self.tl.yview_moveto(a)
        self.tr.yview_moveto(a)

    def _rueda(self, e, paso=None):
        p = paso if paso is not None else (-1 if e.delta > 0 else 1) * 3
        self.tl.yview_scroll(p, 'units')
        self.tr.yview_scroll(p, 'units')
        return 'break'

    def _selsync(self, a, b):
        s = a.selection()
        if s and b.selection() != s:
            b.selection_set(s)
            b.see(s[0])
        self._marcar_pos()

    def pintar_matriz(self):
        m = self.mes_actual()
        if m is None:
            return
        c = self.D.calc_mes(self.serv.get(), self.M())
        sel = self.tr.selection()
        for t in (self.tl, self.tr):
            t.delete(*t.get_children())
        for it in m['items']:
            if it.get('t') == 'c':
                self.tl.insert('', 'end', iid=it['id'], values=('', it['n']), tags=('cat',))
                self.tr.insert('', 'end', iid=it['id'], values=[''] * 39, tags=('cat',))
                continue
            self.tl.insert('', 'end', iid=it['id'], values=(it.get('c', ''), it['n']))
            self.tr.insert('', 'end', iid=it['id'], values=self._valores(it, c[it['id']]), tags=('neg',) if c[it['id']]['sal'] < 0 else ())
        if sel and self.tr.exists(sel[0]):
            self.tr.selection_set(sel)
        self._marcar_pos()

    def _valores(self, it, x):
        d = it.get('d') or []
        ini = fmt(x['ini']) + ('*' if x['iniAuto'] else '') if x['ini'] is not None else ''
        return [ini, fmt(x['rec']) if x['rec'] else '', ', '.join(f[8:10] + '/' + f[5:7] for f in x['fechas']), vto_txt(x['vto'])] + \
               [fmt(d[k]) if k < len(d) and d[k] not in (None, '') else '' for k in range(31)] + [fmt(x['cons']), fmt(x['sal'])]

    def _clic_celda(self, e):
        col = self.tr.identify_column(e.x)
        row = self.tr.identify_row(e.y)
        if col:
            cname = self.tr['columns'][int(col[1:]) - 1]
            self.cur_col = cname
        if row:
            self.tr.selection_set(row)
            self.tl.selection_set(row)
        self.tr.focus_set()
        self._marcar_pos()

    def _marcar_pos(self):
        """Muestra el encabezado de la columna activa resaltado (▼) y la fila/columna en la barra."""
        heads = {'ini': 'EXIST.\nINICIAL', 'rec': 'RECIB.', 'fing': 'F. INGR.', 'vto': 'VENCE', 'tot': 'TOTAL', 'sal': 'SALDO'}
        for cname in self.tr['columns']:
            base = heads.get(cname, cname[1:])
            self.tr.heading(cname, text=('▼ ' + base) if cname == self.cur_col else base)
        s = self.tr.selection()
        if s:
            it = self._item(s[0])
            if it and it.get('t') == 'i':
                donde = f'DÍA {self.cur_col[1:]}' if self.cur_col.startswith('d') else {'ini': 'EXISTENCIA INICIAL'}.get(self.cur_col, self.cur_col.upper())
                self.lb_pos.config(text=f'▶ {it.get("c", "")} {it["n"][:80]}   |   {donde}')
                return
        self.lb_pos.config(text='')

    def _item(self, iid):
        m = self.mes_actual()
        return next((i for i in m['items'] if i['id'] == iid), None) if m else None

    def _tecla(self, e):
        if self.editor:
            return
        if e.keysym in ('Left', 'Right'):
            cols = list(self.tr['columns'])
            k = cols.index(self.cur_col) + (1 if e.keysym == 'Right' else -1)
            if 0 <= k < len(cols):
                self.cur_col = cols[k]
                self._ver_columna()
                self._marcar_pos()
            return 'break'
        if e.keysym == 'Tab':
            self._mover(0, 1)
            return 'break'
        if e.keysym == 'Delete':
            self._poner_valor('')
            return 'break'
        if e.char and e.char.isdigit():
            self.editar_celda(inicial=e.char)
            return 'break'

    def _ver_columna(self):
        cols = list(self.tr['columns'])
        k = cols.index(self.cur_col)
        x = sum(self.tr.column(cname, 'width') for cname in cols[:k])
        total = sum(self.tr.column(cname, 'width') for cname in cols)
        vis = self.tr.winfo_width()
        a, b = self.tr.xview()
        if x < a * total or x + 40 > b * total:
            self.tr.xview_moveto(max(0, (x - vis / 2) / total))

    def editar_celda(self, inicial=None):
        s = self.tr.selection()
        if not s:
            return
        iid = s[0]
        it = self._item(iid)
        if not it or it.get('t') != 'i':
            return
        if self.cur_col == 'rec' or self.cur_col == 'fing' or self.cur_col == 'vto':
            if messagebox.askyesno(APP, 'La cantidad recibida se anota en la pestaña "Ingresos (vales)".\n¿Ir ahora?'):
                self.nb.select(1)
            return
        if self.cur_col not in ['ini'] + self.DIAS:
            return
        self._ver_columna()
        self.update_idletasks()
        bb = self.tr.bbox(iid, self.cur_col)
        if not bb:
            return
        x, y, w, h = bb
        actual = it.get('ini') if self.cur_col == 'ini' else ((it.get('d') or []) + [None] * 31)[int(self.cur_col[1:]) - 1]
        self.editor = tk.Entry(self.tr, justify='center', font=('Segoe UI', 10, 'bold'), bg='#fff9db', relief='solid', bd=1)
        self.editor.place(x=x, y=y, width=max(w, 40), height=h)
        self.editor.insert(0, inicial if inicial is not None else ('' if actual in (None, '') else str(actual)))
        self.editor.focus_set()
        if inicial is None:
            self.editor.select_range(0, 'end')
        self.editor.bind('<Return>', lambda e: self._fin_edicion(1, 0))
        self.editor.bind('<Tab>', lambda e: self._fin_edicion(0, 1) or 'break')
        self.editor.bind('<Shift-Tab>', lambda e: self._fin_edicion(0, -1) or 'break')
        self.editor.bind('<Up>', lambda e: self._fin_edicion(-1, 0))
        self.editor.bind('<Down>', lambda e: self._fin_edicion(1, 0))
        self.editor.bind('<Escape>', lambda e: self._cerrar_editor())
        self.editor.bind('<FocusOut>', lambda e: self._fin_edicion(0, 0, seguir=False))

    def _cerrar_editor(self):
        if self.editor:
            self.editor.destroy()
            self.editor = None
            self.tr.focus_set()

    def _fin_edicion(self, dr, dc, seguir=True):
        if not self.editor:
            return
        txt = self.editor.get().strip()
        self._cerrar_editor()
        if not self._poner_valor(txt):
            return
        if dr or dc:
            self._mover(dr, dc)
            if seguir and dr:
                self.after(10, self.editar_celda)

    def _poner_valor(self, txt):
        s = self.tr.selection()
        if not s:
            return False
        it = self._item(s[0])
        if not it or it.get('t') != 'i':
            return False
        try:
            v = None if txt == '' else eval_cant(txt)
        except ValueError:
            messagebox.showwarning(APP, 'Escriba solo números (puede escribir 10+8).')
            return False
        if self.cur_col == 'ini':
            if v is None:
                it.pop('ini', None)
            else:
                it['ini'] = v
        elif self.cur_col.startswith('d'):
            k = int(self.cur_col[1:]) - 1
            d = it.setdefault('d', [])
            while len(d) <= k:
                d.append(None)
            d[k] = v
            while d and d[-1] is None:
                d.pop()
        else:
            return False
        self.guardar()
        x = self.D.calc_mes(self.serv.get(), self.M())[it['id']]
        self.tr.item(it['id'], values=self._valores(it, x), tags=('neg',) if x['sal'] < 0 else ())
        return True

    def _mover(self, dr, dc):
        cols = list(self.tr['columns'])
        if dc:
            k = cols.index(self.cur_col) + dc
            if 0 <= k < len(cols):
                self.cur_col = cols[k]
                self._ver_columna()
        if dr:
            s = self.tr.selection()
            todos = [i for i in self.tr.get_children() if 'cat' not in self.tr.item(i, 'tags')]
            if s and s[0] in todos:
                k = todos.index(s[0]) + dr
                if 0 <= k < len(todos):
                    self.tr.selection_set(todos[k])
                    self.tl.selection_set(todos[k])
                    self.tr.see(todos[k])
        self._marcar_pos()

    def _editar_nombre(self, e):
        iid = self.tl.identify_row(e.y)
        col = self.tl.identify_column(e.x)
        it = self._item(iid)
        if not it:
            return
        if col == '#1' and it.get('t') == 'i':
            v = simpledialog.askstring(APP, 'Código del insumo (se aplica en todos los meses de este servicio):', initialvalue=it.get('c', ''), parent=self)
            if v is None:
                return
            n0 = nr(it['n'])
            for mid in self.D.meses(self.serv.get()):
                mm = self.D.mes(self.serv.get(), mid)
                for i in mm['items']:
                    if i.get('t') == 'i' and nr(i['n']) == n0:
                        if v.strip():
                            i['c'] = v.strip()
                        else:
                            i.pop('c', None)
                self.D.guardar_mes(self.serv.get(), mid)
        else:
            v = simpledialog.askstring(APP, 'Nombre:', initialvalue=it['n'], parent=self)
            if not v:
                return
            it['n'] = re.sub(r'\s+', ' ', v).strip()
            self.guardar()
        self.pintar_matriz()

    def agregar_grupo(self):
        n = simpledialog.askstring(APP, 'Nombre del grupo (ej.: JERINGAS):', parent=self)
        if n:
            self.mes_actual()['items'].append({'id': nuevo_id('g'), 't': 'c', 'n': n.strip().upper()})
            self.guardar()
            self.pintar_matriz()

    def agregar_insumo(self, grupo=None):
        s = self.tr.selection()
        g = grupo
        if not g and s:
            m = self.mes_actual()
            for i in m['items']:
                if i.get('t') == 'c':
                    g = i['n']
                if i['id'] == s[0]:
                    break
        f = FichaInsumo(self, 'Agregar insumo a la matriz', self.mes_actual(), True, {'grupo': g} if g else None)
        if not f.res:
            return
        r = f.res
        it = self.insertar_insumo(r['c'], r['n'], r['grupo'], self.M() if not r['todos'] else None, r['despues'], r['inicio'], {'nc': r['nc'], 'um': r['um']})
        self.pintar_matriz()
        if it and self.tr.exists(it['id']):
            self.tr.selection_set(it['id'])
            self.tr.see(it['id'])
        self.estado.config(text=f'✔ Insumo agregado: {r["n"][:80]}')

    def insertar_insumo(self, c, n, grupo, solo_mid=None, despues=None, inicio=False, extra=None):
        """Agrega el insumo en el mes (o en este y los siguientes) y en el modelo del servicio."""
        serv = self.serv.get()
        base = {'t': 'i', 'n': n}
        if c:
            base['c'] = c
        for k, v in (extra or {}).items():
            if v:
                base[k] = v
        res = None
        for mid in self.D.meses(serv):
            if (solo_mid and mid != solo_mid) or (not solo_mid and mid < self.M()):
                continue
            m = self.D.mes(serv, mid)
            ya = next((i for i in m['items'] if i.get('t') == 'i' and clave(i) == clave(base)), None)
            if ya:
                res = res or (ya if mid == self.M() else None)
                continue
            it = dict(base, id=nuevo_id(), d=[])
            pos = len(m['items'])
            ref = next((k for k, i in enumerate(m['items']) if despues and i.get('t') == 'i' and clave(i) == despues), None)
            if ref is not None:
                pos = ref + 1
            elif grupo:
                gi = next((k for k, i in enumerate(m['items']) if i.get('t') == 'c' and nr(i['n']) == nr(grupo)), None)
                if gi is None:
                    m['items'].append({'id': nuevo_id('g'), 't': 'c', 'n': grupo})
                    pos = len(m['items'])
                else:
                    pos = gi + 1
                    if not inicio:
                        while pos < len(m['items']) and m['items'][pos].get('t') == 'i':
                            pos += 1
            m['items'].insert(pos, it)
            self.D.guardar_mes(serv, mid)
            if mid == self.M():
                res = it
        mod = self.D.modelo(serv)
        if not any(i.get('t') == 'i' and clave(i) == clave(base) for i in mod['items']):
            pos = len(mod['items'])
            if grupo:
                gi = next((k for k, i in enumerate(mod['items']) if i.get('t') == 'c' and nr(i['n']) == nr(grupo)), None)
                if gi is None:
                    mod['items'].append({'t': 'c', 'n': grupo, 'c': ''})
                    pos = len(mod['items'])
                else:
                    pos = gi + 1
                    while pos < len(mod['items']) and mod['items'][pos].get('t') == 'i':
                        pos += 1
            ref = next((k for k, i in enumerate(mod['items']) if despues and i.get('t') == 'i' and clave(i) == despues), None)
            mod['items'].insert(ref + 1 if ref is not None else pos, dict({'t': 'i', 'n': n, 'c': c}, **{k: v for k, v in (extra or {}).items() if v}))
            self.D.guardar_modelo(serv, mod)
        return res

    def editar_insumo(self):
        s = self.tr.selection()
        it = self._item(s[0]) if s else None
        if not it or it.get('t') != 'i':
            return
        f = FichaInsumo(self, 'Editar insumo', None, False, {'c': it.get('c', ''), 'n': it['n'], 'nc': it.get('nc', ''), 'um': it.get('um', 'C/U')})
        if not f.res:
            return
        r = f.res
        k0 = clave(it)
        for mid in self.D.meses(self.serv.get()):
            if mid < self.M():
                continue
            mm = self.D.mes(self.serv.get(), mid)
            for i in mm['items']:
                if i.get('t') == 'i' and (i is it or clave(i) == k0):
                    i['n'] = r['n']
                    for kk in ('c', 'nc', 'um'):
                        if r[kk]:
                            i[kk] = r[kk]
                        else:
                            i.pop(kk, None)
            self.D.guardar_mes(self.serv.get(), mid)
        self.pintar_matriz()

    def _propagar_orden(self):
        """Copia el orden del mes actual a los meses siguientes y al modelo."""
        if not self.v_orden_todos.get():
            return
        fuente = self.mes_actual()['items']
        idx = {}
        for k, i in enumerate(fuente):
            idx[('c', nr(i['n'])) if i.get('t') == 'c' else ('i', clave(i))] = k

        def ordenar(lista):
            claves, ult = [], -1.0
            for j, i in enumerate(lista):
                kk = ('c', nr(i['n'])) if i.get('t') == 'c' else ('i', clave(i))
                if kk in idx:
                    ult = float(idx[kk])
                    claves.append((ult, j))
                else:
                    claves.append((ult + 0.0001 * (j + 1), j))
            return [lista[j] for _, j in sorted(claves)]
        serv = self.serv.get()
        for mid in self.D.meses(serv):
            if mid > self.M():
                m = self.D.mes(serv, mid)
                m['items'] = ordenar(m['items'])
                self.D.guardar_mes(serv, mid)
        mod = self.D.modelo(serv)
        if mod['items']:
            mod['items'] = ordenar(mod['items'])
            self.D.guardar_modelo(serv, mod)

    def mover_insumo(self, d):
        s = self.tr.selection()
        m = self.mes_actual()
        if not s:
            return
        L = m['items']
        k = next((j for j, i in enumerate(L) if i['id'] == s[0]), None)
        if k is None or not (0 <= k + d < len(L)):
            return
        L[k], L[k + d] = L[k + d], L[k]
        self.guardar()
        self._propagar_orden()
        self.pintar_matriz()
        self.tr.selection_set(s[0])
        self.tl.selection_set(s[0])
        self.tr.see(s[0])

    def mover_a_grupo(self):
        s = self.tr.selection()
        m = self.mes_actual()
        it = self._item(s[0]) if s else None
        if not it or it.get('t') != 'i':
            return
        grupos = [i['n'] for i in m['items'] if i.get('t') == 'c']
        d = ElegirGrupo(self, grupos)
        if not d.res:
            return
        m['items'].remove(it)
        gi = next((j for j, i in enumerate(m['items']) if i.get('t') == 'c' and i['n'] == d.res), None)
        pos = gi + 1
        while pos < len(m['items']) and m['items'][pos].get('t') == 'i':
            pos += 1
        m['items'].insert(pos, it)
        self.guardar()
        self._propagar_orden()
        self.pintar_matriz()
        self.tr.selection_set(it['id'])
        self.tr.see(it['id'])

    def ordenar_grupo(self, por):
        s = self.tr.selection()
        m = self.mes_actual()
        if not s:
            messagebox.showinfo(APP, 'Seleccione un insumo del grupo que quiere ordenar.')
            return
        L = m['items']
        k = next(j for j, i in enumerate(L) if i['id'] == s[0])
        a = k
        while a > 0 and L[a - 1].get('t') == 'i' and L[a].get('t') == 'i':
            a -= 1
        if L[a].get('t') == 'c':
            a += 1
        b = a
        while b < len(L) and L[b].get('t') == 'i':
            b += 1
        L[a:b] = sorted(L[a:b], key=(lambda i: nr(i['n'])) if por == 'n' else (lambda i: (i.get('c') or '99999999', nr(i['n']))))
        self.guardar()
        self._propagar_orden()
        self.pintar_matriz()

    def quitar_insumo(self):
        s = self.tr.selection()
        it = self._item(s[0]) if s else None
        if not it:
            return
        if not messagebox.askyesno(APP, f'¿Quitar "{it["n"][:70]}" de {nombre_mes(self.M())}?\nSe pierde su consumo de este mes.'):
            return
        self.mes_actual()['items'].remove(it)
        self.guardar()
        self.pintar_matriz()

    # ================= PESTAÑA INGRESOS (vales) =================
    def _tab_ingresos(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='🧾 Ingresos (vales)')
        izq = ttk.LabelFrame(f, text='Registrar ingreso del Almacén General (vale)')
        izq.pack(side='left', fill='both', expand=True, padx=4, pady=4)
        fila = ttk.Frame(izq)
        fila.pack(fill='x', pady=4)
        self.v_num, self.v_fecha, self.v_proc = tk.StringVar(), tk.StringVar(value=f'{datetime.date.today():%Y-%m-%d}'), tk.StringVar(value='Almacén General')
        for lab, var, w in (('N° de vale', self.v_num, 14), ('Fecha (AAAA-MM-DD)', self.v_fecha, 12), ('Procedencia', self.v_proc, 18)):
            ttk.Label(fila, text=lab).pack(side='left', padx=(6, 2))
            ttk.Entry(fila, textvariable=var, width=w).pack(side='left')
        f2 = ttk.Frame(izq)
        f2.pack(fill='x', pady=4)
        ttk.Label(f2, text='Insumo (escriba código o nombre):').pack(side='left', padx=4)
        self.v_ins = tk.StringVar()
        self.cb_ins = ttk.Combobox(f2, textvariable=self.v_ins, width=60)
        self.cb_ins.pack(side='left')
        self.cb_ins.bind('<KeyRelease>', self._filtrar_insumos)
        f3 = ttk.Frame(izq)
        f3.pack(fill='x', pady=2)
        self.v_cant, self.v_vto, self.v_lote = tk.StringVar(), tk.StringVar(), tk.StringVar()
        for lab, var, w in (('Cantidad', self.v_cant, 8), ('Vence (AAAA-MM)', self.v_vto, 9), ('Lote / obs.', self.v_lote, 16)):
            ttk.Label(f3, text=lab).pack(side='left', padx=(6, 2))
            ttk.Entry(f3, textvariable=var, width=w).pack(side='left')
        ttk.Button(f3, text='Agregar renglón', command=self._agregar_linea).pack(side='left', padx=6)
        ttk.Button(f3, text='➕ Insumo NUEVO', command=self._insumo_nuevo_vale).pack(side='left')
        self.tv_lin = ttk.Treeview(izq, columns=('c', 'n', 'cant', 'vto', 'obs'), show='headings', height=10)
        for cn, t, w in (('c', 'Código', 70), ('n', 'Insumo', 360), ('cant', 'Cant.', 60), ('vto', 'Vence', 70), ('obs', 'Lote / obs.', 120)):
            self.tv_lin.heading(cn, text=t)
            self.tv_lin.column(cn, width=w, anchor='w' if cn == 'n' else 'center')
        self.tv_lin.pack(fill='both', expand=True, padx=4)
        bb = ttk.Frame(izq)
        bb.pack(fill='x', pady=4)
        ttk.Button(bb, text='Quitar renglón', command=lambda: [self.tv_lin.delete(i) for i in self.tv_lin.selection()]).pack(side='left', padx=4)
        ttk.Button(bb, text='💾 Registrar ingreso y generar recibo', command=self._registrar_vale).pack(side='right', padx=4)
        der = ttk.LabelFrame(f, text='Vales registrados en el mes')
        der.pack(side='left', fill='both', expand=True, padx=4, pady=4)
        self.tv_vales = ttk.Treeview(der, columns=('rec', 'fecha', 'num', 'ren', 'uni'), show='headings')
        for cn, t, w in (('rec', 'Recibo', 70), ('fecha', 'Fecha', 80), ('num', 'N° vale', 90), ('ren', 'Renglones', 70), ('uni', 'Unidades', 80)):
            self.tv_vales.heading(cn, text=t)
            self.tv_vales.column(cn, width=w, anchor='center')
        self.tv_vales.pack(fill='both', expand=True, padx=4)
        b2 = ttk.Frame(der)
        b2.pack(fill='x', pady=4)
        ttk.Button(b2, text='🖨️ Recibo PDF', command=self._recibo_sel).pack(side='left', padx=4)
        ttk.Button(b2, text='Eliminar vale', command=self._borrar_vale).pack(side='left')
        self._mapa_ins = {}

    def _lista_insumos(self):
        m = self.mes_actual()
        self._mapa_ins = {}
        for i in m['items']:
            if i.get('t') == 'i':
                self._mapa_ins[(i.get('c', '') + ' · ' if i.get('c') else '') + i['n']] = i
        return list(self._mapa_ins)

    def _filtrar_insumos(self, e):
        q = nr(self.v_ins.get())
        todos = self._lista_insumos()
        self.cb_ins['values'] = [t for t in todos if q in nr(t)][:60] if q else todos

    def pintar_ingresos(self):
        m = self.mes_actual()
        self.cb_ins['values'] = self._lista_insumos()
        self.tv_vales.delete(*self.tv_vales.get_children())
        for v in sorted(m['vales'], key=lambda v: v.get('fecha', '')):
            self.tv_vales.insert('', 'end', iid=v['id'], values=(v.get('rec') or '—', fecha_txt(v.get('fecha', '')), v.get('num', ''), len(v['lineas']), fmt(sum(l.get('cant') or 0 for l in v['lineas']))))

    def _agregar_linea(self):
        it = self._mapa_ins.get(self.v_ins.get()) or next((i for t, i in self._mapa_ins.items() if self.v_ins.get().strip() and (i.get('c') == self.v_ins.get().strip())), None)
        if not it:
            messagebox.showwarning(APP, 'Elija el insumo de la lista (escriba el código o parte del nombre).')
            return
        try:
            cant = eval_cant(self.v_cant.get())
        except ValueError:
            cant = 0
        if cant <= 0:
            messagebox.showwarning(APP, 'Escriba la cantidad.')
            return
        vto = self.v_vto.get().strip()
        if vto and not re.fullmatch(r'\d{4}-\d{2}', vto):
            messagebox.showwarning(APP, 'El vencimiento va como AAAA-MM (ej.: 2029-05).')
            return
        self.tv_lin.insert('', 'end', values=(it.get('c', ''), it['n'], cant, vto, self.v_lote.get().strip()), tags=(it['id'],))
        self.v_ins.set('')
        self.v_cant.set('')
        self.v_vto.set('')
        self.v_lote.set('')
        self.cb_ins.focus_set()

    def _insumo_nuevo_vale(self):
        f = FichaInsumo(self, 'Insumo NUEVO recibido del Almacén General', self.mes_actual(), False)
        if not f.res:
            return
        r = f.res
        it = self.insertar_insumo(r['c'], r['n'], r['grupo'], None, r['despues'], r['inicio'], {'nc': r['nc'], 'um': r['um']})
        self.pintar_ingresos()
        if it:
            self.v_ins.set((it.get('c', '') + ' · ' if it.get('c') else '') + it['n'])

    def _registrar_vale(self):
        num_v, fecha = self.v_num.get().strip(), self.v_fecha.get().strip()
        if not num_v or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', fecha):
            messagebox.showwarning(APP, 'Escriba el N° de vale y la fecha (AAAA-MM-DD).')
            return
        filas = self.tv_lin.get_children()
        if not filas:
            messagebox.showwarning(APP, 'Agregue al menos un renglón.')
            return
        serv, mid = self.serv.get(), fecha[:7]
        m = self.D.mes(serv, mid)
        if m is None:
            messagebox.showwarning(APP, f'No existe el mes {nombre_mes(mid)} para {serv}. Créelo con "+ Mes siguiente".')
            return
        if any(v.get('num', '').strip().upper() == num_v.upper() for v in m['vales']) and not messagebox.askyesno(APP, 'Ese vale ya está registrado en el mes. ¿Registrarlo otra vez?'):
            return
        lineas = []
        for f in filas:
            vals = self.tv_lin.item(f, 'values')
            iid = self.tv_lin.item(f, 'tags')[0]
            it = next((i for i in m['items'] if i['id'] == iid), None) or next((i for i in m['items'] if i.get('t') == 'i' and clave(i) == clave({'c': vals[0], 'n': vals[1]})), None)
            if not it:
                messagebox.showwarning(APP, f'"{vals[1][:50]}" no está en la matriz de {nombre_mes(mid)}.')
                return
            lineas.append({'item': it['id'], 'cant': int(vals[2]), 'vto': vals[3], 'obs': vals[4]})
        v = {'id': nuevo_id('v'), 'num': num_v, 'fecha': fecha, 'proc': self.v_proc.get().strip(), 'rec': self.D.siguiente_recibo(fecha[:4]), 'lineas': lineas}
        m['vales'].append(v)
        self.guardar(mid)
        self.tv_lin.delete(*filas)
        self.v_num.set('')
        ruta = self.pdf.recibo(serv, mid, v)
        self.pintar_ingresos()
        messagebox.showinfo(APP, f'Ingreso registrado. Recibo N° {v["rec"]}.\nLa matriz y la existencia ya están actualizadas.\n\nRecibo: {ruta}')
        abrir(ruta)

    def _recibo_sel(self):
        s = self.tv_vales.selection()
        if not s:
            return
        m = self.mes_actual()
        v = next(x for x in m['vales'] if x['id'] == s[0])
        if not v.get('rec'):
            v['rec'] = self.D.siguiente_recibo((v.get('fecha') or self.M())[:4])
            self.guardar()
        abrir(self.pdf.recibo(self.serv.get(), self.M(), v))
        self.pintar_ingresos()

    def _borrar_vale(self):
        s = self.tv_vales.selection()
        if s and messagebox.askyesno(APP, '¿Eliminar este vale? Se quita su cantidad recibida de la matriz.'):
            m = self.mes_actual()
            m['vales'] = [v for v in m['vales'] if v['id'] != s[0]]
            self.guardar()
            self.pintar_ingresos()

    # ================= PESTAÑA RETIROS =================
    def _tab_retiros(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='📤 Retiros de bodega')
        a = ttk.LabelFrame(f, text='1. Imprimir hojas de pedido (una por día, con N° correlativo)')
        a.pack(fill='x', padx=4, pady=4)
        self.h_modo = tk.StringVar(value='mes')
        self.h_fecha = tk.StringVar(value=f'{datetime.date.today():%Y-%m-%d}')
        self.h_disp = tk.BooleanVar(value=True)
        ttk.Radiobutton(a, text='Todo el mes (una hoja por día)', variable=self.h_modo, value='mes').pack(side='left', padx=4)
        ttk.Radiobutton(a, text='Solo el día:', variable=self.h_modo, value='dia').pack(side='left')
        ttk.Entry(a, textvariable=self.h_fecha, width=11).pack(side='left')
        ttk.Checkbutton(a, text='Mostrar existencia disponible', variable=self.h_disp).pack(side='left', padx=8)
        ttk.Button(a, text='🖨️ Generar hojas (PDF)', command=self._hojas).pack(side='left', padx=4)
        b = ttk.LabelFrame(f, text='2. Registrar una hoja llenada (descuenta de la existencia)')
        b.pack(fill='both', expand=True, padx=4, pady=4)
        fila = ttk.Frame(b)
        fila.pack(fill='x', pady=3)
        self.r_fecha, self.r_folio, self.r_resp = tk.StringVar(value=f'{datetime.date.today():%Y-%m-%d}'), tk.StringVar(), tk.StringVar()
        self.r_desc = tk.BooleanVar(value=True)
        for lab, var, w in (('Fecha (AAAA-MM-DD)', self.r_fecha, 11), ('N° de hoja', self.r_folio, 8), ('Responsable (quien retira)', self.r_resp, 28)):
            ttk.Label(fila, text=lab).pack(side='left', padx=(6, 2))
            ttk.Entry(fila, textvariable=var, width=w).pack(side='left')
        ttk.Checkbutton(fila, text='Descontar (desmarque si la hoja ya está anotada en la matriz)', variable=self.r_desc).pack(side='left', padx=8)
        ttk.Button(fila, text='Cargar insumos del día', command=self._cargar_hoja).pack(side='left', padx=4)
        self.tv_ret = ttk.Treeview(b, columns=('c', 'n', 'disp', 'ya', 'cant'), show='headings', height=12)
        for cn, t, w in (('c', 'Código', 70), ('n', 'Insumo', 330), ('disp', 'Disponible', 80), ('ya', 'Ya anotado ese día', 120), ('cant', 'CANTIDAD RETIRADA', 140)):
            self.tv_ret.heading(cn, text=t)
            self.tv_ret.column(cn, width=w, anchor='w' if cn == 'n' else 'center')
        self.tv_ret.pack(fill='both', expand=True, padx=4)
        self.tv_ret.bind('<Double-1>', lambda e: self._editar_ret())
        self.tv_ret.bind('<Return>', lambda e: self._editar_ret())
        self.tv_ret.bind('<Key>', lambda e: self._editar_ret(e.char) if e.char and e.char.isdigit() else None)
        bb = ttk.Frame(b)
        bb.pack(fill='x', pady=3)
        ttk.Label(bb, text='Seleccione el insumo y escriba la cantidad (acepta 10+8). Enter baja al siguiente.', foreground='#555').pack(side='left', padx=4)
        ttk.Button(bb, text='💾 Registrar hoja', command=self._registrar_hoja).pack(side='right', padx=4)
        c = ttk.LabelFrame(f, text='3. Retiros registrados en el mes')
        c.pack(fill='both', expand=True, padx=4, pady=4)
        self.tv_rlist = ttk.Treeview(c, columns=('fol', 'fecha', 'resp', 'ins', 'uni', 'desc'), show='headings', height=6)
        for cn, t, w in (('fol', 'N° hoja', 70), ('fecha', 'Fecha', 90), ('resp', 'Responsable', 220), ('ins', 'Insumos', 70), ('uni', 'Unidades', 80), ('desc', 'Descontado', 90)):
            self.tv_rlist.heading(cn, text=t)
            self.tv_rlist.column(cn, width=w, anchor='center')
        self.tv_rlist.pack(side='left', fill='both', expand=True, padx=4)
        bx = ttk.Frame(c)
        bx.pack(side='left', fill='y')
        ttk.Button(bx, text='Eliminar hoja', command=self._borrar_hoja).pack(pady=2)
        ttk.Button(bx, text='🖨️ Control de retiros', command=lambda: abrir(self.pdf.retiros(self.serv.get(), self.M()))).pack(pady=2)
        self._cant_ret = {}

    def pintar_retiros(self):
        m = self.mes_actual()
        self.tv_rlist.delete(*self.tv_rlist.get_children())
        for r in sorted(m['retiros'], key=lambda r: r.get('fecha', '')):
            self.tv_rlist.insert('', 'end', iid=r['id'], values=(r.get('folio', ''), fecha_txt(r.get('fecha', '')), r.get('resp', ''), len(r['lineas']),
                                                                 fmt(sum(l['cant'] for l in r['lineas'])), 'Sí' if r.get('desc', True) else 'No'))
        if not self.tv_ret.get_children():
            self._cargar_hoja()

    def _hojas(self):
        serv = self.serv.get()
        if self.h_modo.get() == 'dia':
            f = self.h_fecha.get().strip()
            if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', f) or self.D.mes(serv, f[:7]) is None:
                messagebox.showwarning(APP, 'Escriba una fecha AAAA-MM-DD de un mes que exista.')
                return
            mid, dias = f[:7], [int(f[8:])]
        else:
            mid = self.M()
            dias = list(range(1, dias_mes(mid) + 1))
        abrir(self.pdf.hojas(serv, mid, dias, self.h_disp.get()))

    def _cargar_hoja(self):
        f = self.r_fecha.get().strip()
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', f):
            return
        serv, mid = self.serv.get(), f[:7]
        m = self.D.mes(serv, mid)
        if m is None:
            messagebox.showwarning(APP, f'No existe el mes {nombre_mes(mid)}.')
            return
        fol = self.D.folio_hoja(serv, f, asignar=False)
        if fol and not self.r_folio.get():
            self.r_folio.set(str(fol))
        c = self.D.calc_mes(serv, mid)
        dia = int(f[8:]) - 1
        self.tv_ret.delete(*self.tv_ret.get_children())
        self._cant_ret = {}
        for it in m['items']:
            if it.get('t') != 'i':
                continue
            x = c[it['id']]
            ya = ((it.get('d') or []) + [None] * 31)[dia]
            self.tv_ret.insert('', 'end', iid=it['id'], values=(it.get('c', ''), nombre_hoja(it), '' if x['sinDatos'] else fmt(x['sal']), fmt(ya) if ya else '', ''))
        self._ret_mid = mid

    def _editar_ret(self, inicial=None):
        s = self.tv_ret.selection()
        if not s:
            return
        iid = s[0]
        bb = self.tv_ret.bbox(iid, 'cant')
        if not bb:
            return
        e = tk.Entry(self.tv_ret, justify='center', font=('Segoe UI', 10, 'bold'), bg='#fff9db')
        e.place(x=bb[0], y=bb[1], width=bb[2], height=bb[3])
        e.insert(0, inicial if inicial is not None else self._cant_ret.get(iid, ''))
        e.focus_set()

        def fin(bajar):
            txt = e.get().strip()
            e.destroy()
            try:
                eval_cant(txt)
            except ValueError:
                messagebox.showwarning(APP, 'Cantidad no válida.')
                return
            self._cant_ret[iid] = txt
            vals = list(self.tv_ret.item(iid, 'values'))
            vals[4] = txt
            self.tv_ret.item(iid, values=vals)
            if bajar:
                hijos = self.tv_ret.get_children()
                k = hijos.index(iid) + 1
                if k < len(hijos):
                    self.tv_ret.selection_set(hijos[k])
                    self.tv_ret.see(hijos[k])
                    self.after(10, self._editar_ret)
            self.tv_ret.focus_set()
        e.bind('<Return>', lambda ev: fin(True))
        e.bind('<Escape>', lambda ev: (e.destroy(), self.tv_ret.focus_set()))
        e.bind('<FocusOut>', lambda ev: fin(False) if e.winfo_exists() else None)

    def _registrar_hoja(self):
        f, resp = self.r_fecha.get().strip(), self.r_resp.get().strip()
        if not resp:
            messagebox.showwarning(APP, 'Escriba el nombre del responsable que retira.')
            return
        serv, mid = self.serv.get(), f[:7]
        m = self.D.mes(serv, mid)
        if m is None or getattr(self, '_ret_mid', None) != mid:
            messagebox.showwarning(APP, 'Pulse "Cargar insumos del día" para la fecha de la hoja.')
            return
        lineas = []
        for iid, txt in self._cant_ret.items():
            v = eval_cant(txt)
            if v > 0:
                lineas.append({'item': iid, 'cant': v})
        if not lineas:
            messagebox.showwarning(APP, 'No escribió ninguna cantidad.')
            return
        desc = self.r_desc.get()
        dia = int(f[8:]) - 1
        ya = [l for l in lineas if ((next(i for i in m['items'] if i['id'] == l['item']).get('d') or []) + [None] * 31)[dia]]
        if desc and ya and not messagebox.askyesno(APP, f'{len(ya)} insumo(s) ya tienen consumo anotado ese día en la matriz.\nSi esta hoja ya estaba anotada se descontaría DOS VECES.\n\n¿Descontar de todos modos?'):
            return
        folio = self.r_folio.get().strip() or str(self.D.folio_hoja(serv, f))
        r = {'id': nuevo_id('rt'), 'folio': folio, 'fecha': f, 'mes': mid, 'serv': serv, 'resp': resp, 'desc': desc, 'lineas': lineas, 'reg': datetime.datetime.now().isoformat(timespec='seconds')}
        if desc:
            self._aplicar_retiro(m, r, 1)
        m['retiros'].append(r)
        self.guardar(mid)
        self._cant_ret = {}
        self.r_folio.set('')
        self.r_resp.set('')
        self._cargar_hoja()
        self.pintar_retiros()
        messagebox.showinfo(APP, f'Hoja N° {folio} registrada: {len(lineas)} insumos, {sum(l["cant"] for l in lineas)} unidades {"descontadas" if desc else "registradas sin descontar"}.')

    def _aplicar_retiro(self, m, r, signo):
        if not r.get('desc', True):
            return
        d0 = int(r['fecha'][8:]) - 1
        for l in r['lineas']:
            it = next((i for i in m['items'] if i['id'] == l['item']), None)
            if not it:
                continue
            d = it.setdefault('d', [])
            while len(d) <= d0:
                d.append(None)
            v = (d[d0] or 0) + signo * l['cant']
            d[d0] = None if v <= 0 else v
            while d and d[-1] is None:
                d.pop()

    def _borrar_hoja(self):
        s = self.tv_rlist.selection()
        if not s:
            return
        m = self.mes_actual()
        r = next(x for x in m['retiros'] if x['id'] == s[0])
        if messagebox.askyesno(APP, f'¿Eliminar la hoja N° {r.get("folio")} de {r.get("resp")}? Las cantidades vuelven a la existencia.'):
            self._aplicar_retiro(m, r, -1)
            m['retiros'].remove(r)
            self.guardar()
            self.pintar_retiros()

    # ================= PESTAÑA EXISTENCIAS =================
    def _tab_existencias(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='📊 Existencias')
        top = ttk.Frame(f)
        top.pack(fill='x', pady=4)
        self.e_filtro = tk.StringVar(value='Con existencia o movimiento')
        ttk.Label(top, text='Mostrar:').pack(side='left', padx=4)
        cb = ttk.Combobox(top, textvariable=self.e_filtro, state='readonly', width=34,
                          values=['Con existencia o movimiento', 'Todos', 'Saldo negativo', 'Agotados', 'Falta existencia inicial'])
        cb.pack(side='left')
        cb.bind('<<ComboboxSelected>>', lambda e: self.pintar_existencias())
        ttk.Button(top, text='🖨️ PDF', command=lambda: abrir(self.pdf.existencias(self.serv.get(), self.M()))).pack(side='left', padx=6)
        self.lb_ex = ttk.Label(top, text='', foreground='#1f3a68', font=('Segoe UI', 9, 'bold'))
        self.lb_ex.pack(side='left', padx=10)
        self.tv_ex = ttk.Treeview(f, columns=('c', 'n', 'ant', 'ini', 'rec', 'cons', 'sal', 'est'), show='headings')
        for cn, t, w in (('c', 'Código', 70), ('n', 'Insumo', 420), ('ant', 'Saldo mes anterior', 110), ('ini', 'Existencia inicial', 120), ('rec', 'Ingresos', 70),
                         ('cons', 'Consumo', 70), ('sal', 'Saldo al cierre', 100), ('est', 'Estado', 160)):
            self.tv_ex.heading(cn, text=t)
            self.tv_ex.column(cn, width=w, anchor='w' if cn == 'n' else 'center')
        self.tv_ex.tag_configure('neg', foreground='#c00000')
        self.tv_ex.tag_configure('ag', foreground='#a04000')
        sb = ttk.Scrollbar(f, orient='vertical', command=self.tv_ex.yview)
        self.tv_ex.configure(yscrollcommand=sb.set)
        self.tv_ex.pack(side='left', fill='both', expand=True, padx=4)
        sb.pack(side='left', fill='y')

    def pintar_existencias(self):
        m = self.mes_actual()
        c = self.D.calc_mes(self.serv.get(), self.M())
        self.tv_ex.delete(*self.tv_ex.get_children())
        cont = {'neg': 0, 'ag': 0, 'ok': 0, 'inc': 0}
        fil = self.e_filtro.get()
        for it in m['items']:
            if it.get('t') != 'i':
                continue
            x = c[it['id']]
            est = 'Sin datos' if x['sinDatos'] else ('Negativo' if x['sal'] < 0 else ('Falta existencia inicial' if x['incompleto'] else ('Agotado' if x['sal'] == 0 else 'Con existencia')))
            cont['neg' if est == 'Negativo' else 'ag' if est == 'Agotado' else 'inc' if est.startswith('Falta') else 'ok'] += 0 if x['sinDatos'] else 1
            if (fil == 'Con existencia o movimiento' and x['sinDatos']) or (fil == 'Saldo negativo' and est != 'Negativo') or (fil == 'Agotados' and est != 'Agotado') or (fil == 'Falta existencia inicial' and not est.startswith('Falta')):
                continue
            ini = '' if x['ini'] is None else (fmt(x['ini']) + (' (auto)' if x['iniAuto'] else ' (conteo)') + (f' ajuste {x["ajuste"]:+}' if x['ajuste'] else ''))
            self.tv_ex.insert('', 'end', values=(it.get('c', ''), it['n'], '' if x['ant'] is None else fmt(x['ant']), ini, fmt(x['rec']) if x['rec'] else '', fmt(x['cons']) if x['cons'] else '',
                                                 '' if x['sinDatos'] else fmt(x['sal']), est), tags=('neg',) if est == 'Negativo' else ('ag',) if est == 'Agotado' else ())
        self.lb_ex.config(text=f'Con existencia: {cont["ok"]}   ·   Agotados: {cont["ag"]}   ·   Negativos: {cont["neg"]}   ·   Falta existencia inicial: {cont["inc"]}')

    # ================= PESTAÑA MATRIZ POR PERÍODO =================
    def _tab_periodo(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='🗓️ Matriz por período')
        top = ttk.Frame(f)
        top.pack(fill='x', pady=4)
        self.p_desde, self.p_hasta = tk.StringVar(), tk.StringVar()
        ttk.Label(top, text='Desde:').pack(side='left', padx=4)
        self.cb_pd = ttk.Combobox(top, textvariable=self.p_desde, state='readonly', width=12)
        self.cb_pd.pack(side='left')
        ttk.Label(top, text='Hasta:').pack(side='left', padx=4)
        self.cb_ph = ttk.Combobox(top, textvariable=self.p_hasta, state='readonly', width=12)
        self.cb_ph.pack(side='left')
        ttk.Button(top, text='Ver', command=self._ver_periodo).pack(side='left', padx=4)
        ttk.Button(top, text='💾 Guardar matriz del período (JSON + PDF)', command=self._guardar_periodo).pack(side='left', padx=4)
        self.tv_per = ttk.Treeview(f, show='headings')
        sb = ttk.Scrollbar(f, orient='vertical', command=self.tv_per.yview)
        hb = ttk.Scrollbar(f, orient='horizontal', command=self.tv_per.xview)
        self.tv_per.configure(yscrollcommand=sb.set, xscrollcommand=hb.set)
        hb.pack(side='bottom', fill='x')
        self.tv_per.pack(side='left', fill='both', expand=True, padx=4)
        sb.pack(side='left', fill='y')

    def pintar_periodo(self):
        ms = self.D.meses(self.serv.get())
        self.cb_pd['values'] = ms
        self.cb_ph['values'] = ms
        if self.p_desde.get() not in ms:
            self.p_desde.set(ms[0] if ms else '')
        if self.p_hasta.get() not in ms:
            self.p_hasta.set(self.M() or (ms[-1] if ms else ''))
        self._ver_periodo()

    def consolidado(self, serv, desde, hasta):
        meses = [x for x in self.D.meses(serv) if desde <= x <= hasta]
        todo = self.D.calcular(serv, hasta)
        P, orden = {}, []
        for mid in meses:
            m = self.D.mes(serv, mid)
            for it in m['items']:
                if it.get('t') != 'i':
                    continue
                x = todo[mid][it['id']]
                k = x['clave']
                if k not in P:
                    P[k] = {'c': it.get('c', ''), 'n': it['n'], 'ini': x['ini'], 'rec': 0, 'cons': 0, 'sal': 0, 'cons_mes': {}}
                    orden.append(k)
                o = P[k]
                o['c'] = it.get('c', '') or o['c']
                o['n'] = it['n']
                o['rec'] += x['rec']
                o['cons'] += x['cons']
                o['cons_mes'][mid] = x['cons']
                if not x['sinDatos']:
                    o['sal'] = x['sal']
        return [P[k] for k in orden if P[k]['rec'] or P[k]['cons'] or P[k]['ini'] or P[k]['sal']], meses

    def _ver_periodo(self):
        d, h = self.p_desde.get(), self.p_hasta.get()
        if not d or not h:
            return
        if d > h:
            d, h = h, d
        filas, meses = self.consolidado(self.serv.get(), d, h)
        cols = ['c', 'n', 'ini'] + meses + ['rec', 'cons', 'sal']
        self.tv_per.delete(*self.tv_per.get_children())
        self.tv_per['columns'] = cols
        for cn in cols:
            t = {'c': 'Código', 'n': 'Insumo', 'ini': 'Exist. inicial', 'rec': 'Ingresos', 'cons': 'Consumo total', 'sal': 'Saldo final'}.get(cn, nombre_mes(cn)[:3] + ' ' + cn[:4] if '-' in cn else cn)
            self.tv_per.heading(cn, text=t)
            self.tv_per.column(cn, width=380 if cn == 'n' else 90, anchor='w' if cn == 'n' else 'center', stretch=False)
        for o in filas:
            self.tv_per.insert('', 'end', values=[o['c'], o['n'], '—' if o['ini'] is None else fmt(o['ini'])] + [fmt(o['cons_mes'].get(x, '')) for x in meses] + [fmt(o['rec']), fmt(o['cons']), fmt(o['sal'])])

    def _guardar_periodo(self):
        d, h = sorted([self.p_desde.get(), self.p_hasta.get()])
        serv = self.serv.get()
        filas, meses = self.consolidado(serv, d, h)
        ruta = os.path.join(self.D.base, 'matrices', f'Matriz {d} a {h} {serv}.json')
        self.D._escribir(ruta, {'servicio': serv, 'desde': d, 'hasta': h, 'meses': meses, 'generado': datetime.datetime.now().isoformat(timespec='seconds'), 'insumos': filas})
        pdf = self.pdf.periodo(serv, d, h)
        messagebox.showinfo(APP, f'Matriz del período guardada:\n{ruta}\n{pdf}')
        abrir(pdf)

    # ================= PESTAÑA MODELOS =================
    def _tab_modelos(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='📋 Modelo de insumos')
        top = ttk.Frame(f)
        top.pack(fill='x', pady=4)
        ttk.Label(top, text='Lista modelo del servicio: se usa al crear meses nuevos. ', foreground='#555').pack(side='left', padx=4)
        for t, cmd in (('+ Grupo', self._mod_grupo), ('+ Insumo', self._mod_insumo), ('Editar', self._mod_editar), ('▲ Subir', lambda: self._mod_mover(-1)),
                       ('▼ Bajar', lambda: self._mod_mover(1)), ('Quitar', self._mod_quitar), ('Copiar del mes actual', self._mod_desde_mes), ('Aplicar al mes actual', self._mod_aplicar)):
            ttk.Button(top, text=t, command=cmd).pack(side='left', padx=2)
        self.tv_mod = ttk.Treeview(f, columns=('t', 'c', 'n'), show='headings')
        for cn, t, w in (('t', 'Tipo', 70), ('c', 'Código', 90), ('n', 'Nombre', 800)):
            self.tv_mod.heading(cn, text=t)
            self.tv_mod.column(cn, width=w, anchor='w')
        self.tv_mod.tag_configure('cat', background='#d9e1f2', font=('Segoe UI', 9, 'bold'))
        sb = ttk.Scrollbar(f, orient='vertical', command=self.tv_mod.yview)
        self.tv_mod.configure(yscrollcommand=sb.set)
        self.tv_mod.pack(side='left', fill='both', expand=True, padx=4)
        sb.pack(side='left', fill='y')

    def pintar_modelo(self):
        self._mod = self.D.modelo(self.serv.get())
        self.tv_mod.delete(*self.tv_mod.get_children())
        for k, i in enumerate(self._mod['items']):
            self.tv_mod.insert('', 'end', iid=str(k), values=('GRUPO' if i['t'] == 'c' else 'insumo', i.get('c', ''), i['n']), tags=('cat',) if i['t'] == 'c' else ())

    def _mod_sel(self):
        s = self.tv_mod.selection()
        return int(s[0]) if s else None

    def _mod_guardar(self, sel=None):
        self.D.guardar_modelo(self.serv.get(), self._mod)
        self.pintar_modelo()
        if sel is not None and self.tv_mod.exists(str(sel)):
            self.tv_mod.selection_set(str(sel))
            self.tv_mod.see(str(sel))

    def _mod_grupo(self):
        n = simpledialog.askstring(APP, 'Nombre del grupo:', parent=self)
        if n:
            k = self._mod_sel()
            pos = len(self._mod['items']) if k is None else k + 1
            self._mod['items'].insert(pos, {'t': 'c', 'n': n.strip().upper(), 'c': ''})
            self._mod_guardar(pos)

    def _mod_insumo(self):
        f = FichaInsumo(self, 'Insumo del modelo', {'items': [dict(i, id=str(k)) for k, i in enumerate(self._mod['items'])]}, False)
        if f.res:
            r = f.res
            k = self._mod_sel()
            pos = len(self._mod['items']) if k is None else k + 1
            if r['despues']:
                pos = next((j + 1 for j, i in enumerate(self._mod['items']) if i.get('t') == 'i' and clave(i) == r['despues']), pos)
            elif r['grupo']:
                gi = next((j for j, i in enumerate(self._mod['items']) if i['t'] == 'c' and nr(i['n']) == nr(r['grupo'])), None)
                if gi is not None:
                    pos = gi + 1
                    if not r['inicio']:
                        while pos < len(self._mod['items']) and self._mod['items'][pos]['t'] == 'i':
                            pos += 1
            self._mod['items'].insert(pos, {k2: v for k2, v in {'t': 'i', 'n': r['n'], 'c': r['c'], 'nc': r['nc'], 'um': r['um']}.items() if v or k2 in ('t', 'n', 'c')})
            self._mod_guardar(pos)

    def _mod_editar(self):
        k = self._mod_sel()
        if k is None:
            return
        i = self._mod['items'][k]
        n = simpledialog.askstring(APP, 'Nombre:', initialvalue=i['n'], parent=self)
        if not n:
            return
        i['n'] = n.strip()
        if i['t'] == 'i':
            c = simpledialog.askstring(APP, 'Código:', initialvalue=i.get('c', ''), parent=self)
            if c is not None:
                i['c'] = c.strip()
        self._mod_guardar(k)

    def _mod_mover(self, d):
        k = self._mod_sel()
        if k is None or not (0 <= k + d < len(self._mod['items'])):
            return
        L = self._mod['items']
        L[k], L[k + d] = L[k + d], L[k]
        self._mod_guardar(k + d)

    def _mod_quitar(self):
        k = self._mod_sel()
        if k is not None and messagebox.askyesno(APP, '¿Quitar del modelo? (No borra datos de los meses ya creados)'):
            del self._mod['items'][k]
            self._mod_guardar()

    def _mod_desde_mes(self):
        if messagebox.askyesno(APP, f'¿Reemplazar el modelo con la lista de insumos de {nombre_mes(self.M())}?'):
            self._mod = {'servicio': self.serv.get(), 'items': [{'t': i['t'], 'n': i['n'], 'c': i.get('c', '')} for i in self.mes_actual()['items']]}
            self._mod_guardar()

    def _mod_aplicar(self):
        m = self.mes_actual()
        claves = {clave(i) for i in m['items'] if i.get('t') == 'i'}
        grupo, n = '', 0
        for i in self._mod['items']:
            if i['t'] == 'c':
                grupo = i['n']
                continue
            if clave(i) not in claves:
                self.insertar_insumo(i.get('c', ''), i['n'], grupo, self.M())
                n += 1
        messagebox.showinfo(APP, f'Se agregaron {n} insumo(s) del modelo a {nombre_mes(self.M())}.')

    # ================= PESTAÑA DOCUMENTOS =================
    def _tab_documentos(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='🖨️ Documentos')
        ttk.Label(f, text='Los documentos se guardan en PDF en Documentos\\BodegaEmergencias\\documentos\\(mes y servicio) y se abren para imprimir.',
                  foreground='#555').pack(anchor='w', padx=8, pady=6)
        g = ttk.Frame(f)
        g.pack(anchor='w', padx=8)
        docs = [('Matriz de entradas y salidas del mes (M03, oficio horizontal)', lambda: self.pdf.matriz(self.serv.get(), self.M())),
                ('Existencias del mes (de un mes a otro)', lambda: self.pdf.existencias(self.serv.get(), self.M())),
                ('Tarjetas de control de existencia (kardex)', lambda: self.pdf.kardex(self.serv.get(), self.M())),
                ('Control de retiros del mes', lambda: self.pdf.retiros(self.serv.get(), self.M())),
                ('Informe mensual de control antisoborno', self._doc_antisoborno),
                ('Informe para auditoría interna y externa (por período)', self._doc_auditoria),
                ('Matriz por período (consolidada)', lambda: self.nb.select(4))]
        for k, (t, cmd) in enumerate(docs):
            ttk.Button(g, text='🖨️  ' + t, width=70, command=lambda c=cmd: self._gen(c)).grid(row=k, column=0, sticky='w', pady=3)
        ttk.Button(g, text='📁  Generar TODOS los documentos del mes', width=70, command=self._todos).grid(row=len(docs), column=0, sticky='w', pady=12)

    def _gen(self, cmd):
        try:
            self.config(cursor='watch')
            self.update()
            r = cmd()
            if r:
                abrir(r)
                self.estado.config(text='✔ Documento: ' + r)
        except Exception as e:
            messagebox.showerror(APP, f'No se pudo generar el documento:\n{e}')
        finally:
            self.config(cursor='')

    def _todos(self):
        serv, mid = self.serv.get(), self.M()
        try:
            for fn in (self.pdf.matriz, self.pdf.existencias, self.pdf.kardex, self.pdf.retiros, self.pdf.antisoborno):
                fn(serv, mid)
            abrir(self.pdf.carpeta(serv, mid))
        except Exception as e:
            messagebox.showerror(APP, str(e))

    def _doc_antisoborno(self):
        m = self.mes_actual()
        a = m.setdefault('antisoborno', {})
        d = DialogoTexto(self, 'Informe antisoborno: revisión del responsable', a)
        if d.res is None:
            return None
        m['antisoborno'] = d.res
        self.guardar()
        return self.pdf.antisoborno(self.serv.get(), self.M())

    def _doc_auditoria(self):
        ms = self.D.meses(self.serv.get())
        d = simpledialog.askstring(APP, f'Desde (AAAA-MM). Meses disponibles: {", ".join(ms)}', initialvalue=ms[0] if ms else '', parent=self)
        if not d:
            return None
        h = simpledialog.askstring(APP, 'Hasta (AAAA-MM):', initialvalue=self.M(), parent=self)
        if not h:
            return None
        return self.pdf.auditoria(self.serv.get(), min(d, h), max(d, h))

    def catalogo(self):
        """Insumos conocidos en todos los servicios, meses y modelos (para buscar al agregar)."""
        vistos, out = set(), []
        for serv in self.D.servicios():
            fuentes = [self.D.modelo(serv)['items']] + [self.D.mes(serv, mid)['items'] for mid in reversed(self.D.meses(serv))]
            for items in fuentes:
                g = ''
                for i in items:
                    if i.get('t') == 'c':
                        g = i['n']
                        continue
                    k = clave(i)
                    if k in vistos:
                        continue
                    vistos.add(k)
                    out.append({'c': i.get('c', ''), 'n': i['n'], 'nc': i.get('nc', ''), 'um': i.get('um', ''), 'grupo': g})
        return sorted(out, key=lambda x: nr(x['n']))

    def historia(self, serv, hasta):
        """Movimientos por insumo (para el kardex): existencia inicial, ingresos por vale y salida mensual."""
        H, saldo = {}, {}
        for mid in self.D.meses(serv):
            if mid > hasta:
                break
            m = self.D.mes(serv, mid)
            for it in m['items']:
                if it.get('t') != 'i':
                    continue
                k = clave(it)
                L = H.setdefault(k, [])
                s = saldo.get(k)
                ini = num(it.get('ini'))
                if s is not None and s < 0 and ini is None:
                    L.append({'f': mid + '-01', 'con': f'Saldo anterior negativo ({fmt(s)}): inicia en 0', 'saldo': 0})
                    s = 0
                if ini is not None and (s is None or ini != s):
                    L.append({'f': mid + '-01', 'con': f'Existencia inicial de {nombre_mes(mid)}' + (f' (ajuste {ini - s:+})' if s is not None else ''), 'saldo': ini})
                    s = ini
                for v in sorted(m['vales'], key=lambda v: v.get('fecha', '')):
                    for l in v['lineas']:
                        if l['item'] == it['id']:
                            s = (s or 0) + (l.get('cant') or 0)
                            L.append({'f': v.get('fecha', ''), 'con': f'Ingreso vale N° {v.get("num", "")}', 'proc': 'Almacén General', 'lote': l.get('obs', ''), 'vto': l.get('vto', ''), 'ent': l.get('cant'), 'saldo': s})
                cons = sum(x or 0 for x in (it.get('d') or []) if isinstance(x, (int, float)))
                if cons:
                    s = (s or 0) - cons
                    L.append({'f': f'{mid}-{dias_mes(mid):02d}', 'con': f'Salida: consumo de {nombre_mes(mid)} (matriz)', 'sal': cons, 'saldo': s})
                if s is not None:
                    saldo[k] = s
        return H

    # ================= PESTAÑA AJUSTES / RESPALDO =================
    def _tab_ajustes(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text='⚙️ Respaldo y ajustes')
        a = ttk.LabelFrame(f, text='Encabezado de los documentos')
        a.pack(fill='x', padx=6, pady=6)
        self.cfg_vars = {}
        for k, (key, lab) in enumerate([('l1', 'Línea 1'), ('l2', 'Línea 2'), ('tit', 'Título'), ('cod', 'Código'), ('ver', 'Versión'), ('vig', 'Vigencia'),
                                        ('resp', 'Responsable'), ('f1', 'Firma 1'), ('f2', 'Firma 2'), ('f3', 'Firma 3')]):
            ttk.Label(a, text=lab).grid(row=k if k < 7 else k + 5, column=0, sticky='w', padx=4, pady=1)
            v = tk.StringVar()
            ttk.Entry(a, textvariable=v, width=110).grid(row=k if k < 7 else k + 5, column=1, sticky='w', pady=1)
            self.cfg_vars[key] = v
        self.v_firmas = tk.BooleanVar(value=bool(self.D.cfg.get('firmas')))
        ttk.Checkbutton(a, text='Incluir firmas al pie de los documentos (Elaborado por / Revisado por / Vo. Bo.) — opcional', variable=self.v_firmas).grid(row=10, column=1, sticky='w', pady=2)
        ttk.Button(a, text='Guardar encabezado', command=self._guardar_cfg).grid(row=11, column=1, sticky='w', pady=4)
        b = ttk.LabelFrame(f, text='Datos y respaldo')
        b.pack(fill='x', padx=6, pady=6)
        ttk.Button(b, text='📥 Importar respaldo del programa anterior (.json)', command=self.importar).pack(side='left', padx=4, pady=6)
        ttk.Button(b, text='💾 Crear respaldo (ZIP)', command=lambda: (messagebox.showinfo(APP, 'Respaldo creado:\n' + self.D.respaldo_zip()))).pack(side='left', padx=4)
        ttk.Button(b, text='📂 Abrir carpeta de datos', command=lambda: abrir(self.D.base)).pack(side='left', padx=4)
        ttk.Label(f, text=f'Carpeta de datos: {self.D.base}\nCada mes y servicio es un archivo JSON en "datos". Para pasar los datos a otra PC copie la carpeta completa '
                          '(o use el respaldo ZIP) en una memoria USB.', foreground='#555', justify='left').pack(anchor='w', padx=8)

    def pintar_ajustes(self):
        for k, v in self.cfg_vars.items():
            v.set(self.D.cfg.get(k, ''))
        self.v_firmas.set(bool(self.D.cfg.get('firmas')))

    def _guardar_cfg(self):
        for k, v in self.cfg_vars.items():
            self.D.cfg[k] = v.get()
        self.D.cfg['firmas'] = bool(self.v_firmas.get())
        self.D.guardar_cfg()
        messagebox.showinfo(APP, 'Encabezado guardado.')

    def importar(self):
        p = filedialog.askopenfilename(title='Respaldo del programa anterior', filetypes=[('Respaldo JSON', '*.json'), ('Todos', '*.*')])
        if not p:
            return
        try:
            obj = json.load(open(p, encoding='utf-8'))
            serv = self.serv.get() or 'EMERGENCIA'
            if self.D.meses(serv) and not messagebox.askyesno(APP, f'Ya hay meses de {serv}. Los meses que vengan en el respaldo se REEMPLAZAN. ¿Continuar?'):
                return
            self.D.respaldo_zip()
            n = self.D.importar_respaldo(obj, serv)
            messagebox.showinfo(APP, f'Se importaron {n} mes(es) de {serv}.')
            self.D.cache.clear()
            self.cb_serv['values'] = self.D.servicios()
            self._cambio_serv()
        except Exception as e:
            messagebox.showerror(APP, f'No se pudo importar:\n{e}')


def nr_txt(s):
    s = unicodedata.normalize('NFD', str(s or ''))
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn').upper()


UNIDADES = ['C/U', 'Caja', 'Rollo', 'Par', 'Paquete', 'Frasco', 'Litro', 'Galón', 'Resma', 'Tubo', 'Bolsa', 'Sobre']


class FichaInsumo(tk.Toplevel):
    """Ventana profesional para agregar o editar un insumo."""
    def __init__(self, app, titulo, mes=None, con_todos=True, inicial=None):
        super().__init__(app)
        self.app, self.res, self.mes = app, None, mes
        self.title(titulo)
        self.transient(app)
        self.resizable(False, False)
        self.configure(bg='#f4f6fa')
        ini = inicial or {}
        cab = tk.Frame(self, bg='#1f3a68')
        cab.pack(fill='x')
        tk.Label(cab, text='➕  ' + titulo, bg='#1f3a68', fg='white', font=('Segoe UI', 13, 'bold')).pack(anchor='w', padx=16, pady=(12, 0))
        tk.Label(cab, text='Busque en el catálogo o escriba los datos. Elija el grupo y en qué lugar de la matriz va.', bg='#1f3a68', fg='#c9d6ee',
                 font=('Segoe UI', 9)).pack(anchor='w', padx=16, pady=(0, 12))
        cuerpo = tk.Frame(self, bg='#f4f6fa')
        cuerpo.pack(fill='both', padx=16, pady=10)
        # 1. catálogo
        s1 = ttk.LabelFrame(cuerpo, text=' 1. Buscar en el catálogo (insumos ya usados en cualquier mes o servicio) ')
        s1.grid(row=0, column=0, sticky='nsew', pady=(0, 8))
        self.v_bus = tk.StringVar()
        e = ttk.Entry(s1, textvariable=self.v_bus, width=92)
        e.pack(fill='x', padx=8, pady=(6, 4))
        self.lst = tk.Listbox(s1, height=6, font=('Segoe UI', 9), activestyle='none', selectbackground='#cfe0ff', selectforeground='black', relief='flat', bd=1)
        self.lst.pack(fill='x', padx=8, pady=(0, 8))
        self.cat = app.catalogo()
        self.v_bus.trace_add('write', lambda *a: self._filtrar())
        self.lst.bind('<<ListboxSelect>>', lambda e: self._elegir())
        # 2. datos
        s2 = ttk.LabelFrame(cuerpo, text=' 2. Datos del insumo ')
        s2.grid(row=1, column=0, sticky='nsew', pady=(0, 8))
        self.v_c, self.v_nc, self.v_um = tk.StringVar(value=ini.get('c', '')), tk.StringVar(value=ini.get('nc', '')), tk.StringVar(value=ini.get('um', 'C/U'))
        ttk.Label(s2, text='Código institucional').grid(row=0, column=0, sticky='w', padx=8, pady=(6, 2))
        ttk.Entry(s2, textvariable=self.v_c, width=16, font=('Segoe UI', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=8)
        ttk.Label(s2, text='Unidad de medida').grid(row=0, column=1, sticky='w', padx=8, pady=(6, 2))
        ttk.Combobox(s2, textvariable=self.v_um, values=UNIDADES, width=12).grid(row=1, column=1, sticky='w', padx=8)
        ttk.Label(s2, text='Nombre oficial (como en el Almacén General)').grid(row=2, column=0, columnspan=3, sticky='w', padx=8, pady=(8, 2))
        self.t_n = tk.Text(s2, width=78, height=3, font=('Segoe UI', 10), wrap='word', relief='solid', bd=1)
        self.t_n.grid(row=3, column=0, columnspan=3, sticky='w', padx=8)
        self.t_n.insert('1.0', ini.get('n', ''))
        ttk.Label(s2, text='Nombre corto para la hoja de pedido (opcional)').grid(row=4, column=0, columnspan=3, sticky='w', padx=8, pady=(8, 2))
        ttk.Entry(s2, textvariable=self.v_nc, width=50).grid(row=5, column=0, columnspan=3, sticky='w', padx=8, pady=(0, 8))
        # 3. ubicación
        s3 = ttk.LabelFrame(cuerpo, text=' 3. Ubicación en la matriz ')
        s3.grid(row=2, column=0, sticky='nsew')
        self.grupos = [i['n'] for i in (mes or {'items': []})['items'] if i.get('t') == 'c']
        self.v_g, self.v_pos = tk.StringVar(value=ini.get('grupo', self.grupos[0] if self.grupos else '')), tk.StringVar(value='Al final del grupo')
        ttk.Label(s3, text='Grupo').grid(row=0, column=0, sticky='w', padx=8, pady=(6, 2))
        self.cb_g = ttk.Combobox(s3, textvariable=self.v_g, values=self.grupos + ['➕ Grupo nuevo…'], width=40, state='readonly' if self.grupos else 'normal')
        self.cb_g.grid(row=1, column=0, sticky='w', padx=8)
        self.cb_g.bind('<<ComboboxSelected>>', lambda e: self._grupo())
        ttk.Label(s3, text='Posición').grid(row=0, column=1, sticky='w', padx=8, pady=(6, 2))
        self.cb_pos = ttk.Combobox(s3, textvariable=self.v_pos, width=52, state='readonly')
        self.cb_pos.grid(row=1, column=1, sticky='w', padx=8)
        self.v_todos = tk.BooleanVar(value=True)
        if con_todos:
            ttk.Checkbutton(s3, text='Agregar también en los meses siguientes y en el modelo del servicio', variable=self.v_todos).grid(row=2, column=0, columnspan=2, sticky='w', padx=8, pady=8)
        else:
            ttk.Label(s3, text='').grid(row=2, column=0, pady=4)
        self.msg = tk.Label(cuerpo, text='', bg='#f4f6fa', fg='#c00000', font=('Segoe UI', 9, 'bold'))
        self.msg.grid(row=3, column=0, sticky='w', pady=(6, 0))
        pie = tk.Frame(self, bg='#e6ebf3')
        pie.pack(fill='x')
        tk.Button(pie, text='Cancelar', command=self.destroy, relief='flat', bg='#d5dbe5', padx=16, pady=6, font=('Segoe UI', 10)).pack(side='right', padx=(6, 16), pady=10)
        tk.Button(pie, text='💾  Guardar insumo', command=self._ok, relief='flat', bg='#1a7f37', fg='white', activebackground='#15692d', padx=18, pady=6,
                  font=('Segoe UI', 10, 'bold')).pack(side='right', pady=10)
        self._filtrar()
        self._grupo()
        self.bind('<Escape>', lambda e: self.destroy())
        self.grab_set()
        e.focus_set()
        self.update_idletasks()
        self.geometry(f'+{app.winfo_rootx() + 160}+{app.winfo_rooty() + 60}')
        self.wait_window()

    def _filtrar(self):
        q = nr(self.v_bus.get())
        self.lst.delete(0, 'end')
        self._vis = [c for c in self.cat if not q or q in nr((c.get('c') or '') + ' ' + c['n'])][:80]
        for c in self._vis:
            self.lst.insert('end', f"{c.get('c') or '—':>10}   {c['n'][:110]}")

    def _elegir(self):
        s = self.lst.curselection()
        if not s:
            return
        c = self._vis[s[0]]
        self.v_c.set(c.get('c', ''))
        self.t_n.delete('1.0', 'end')
        self.t_n.insert('1.0', c['n'])
        if c.get('nc'):
            self.v_nc.set(c['nc'])
        if c.get('um'):
            self.v_um.set(c['um'])
        if c.get('grupo') and c['grupo'] in self.grupos:
            self.v_g.set(c['grupo'])
            self._grupo()

    def _grupo(self):
        if self.v_g.get() == '➕ Grupo nuevo…':
            g = simpledialog.askstring('Grupo nuevo', 'Nombre del grupo nuevo (ej.: TUBOS ENDOTRAQUEALES):', parent=self)
            if g:
                g = g.strip().upper()
                self.grupos.append(g)
                self.cb_g['values'] = self.grupos + ['➕ Grupo nuevo…']
                self.v_g.set(g)
            else:
                self.v_g.set(self.grupos[0] if self.grupos else '')
        opciones = ['Al final del grupo', 'Al inicio del grupo']
        self._despues = {}
        if self.mes:
            g, dentro = None, False
            for i in self.mes['items']:
                if i.get('t') == 'c':
                    dentro = nr(i['n']) == nr(self.v_g.get())
                    continue
                if dentro:
                    etq = 'Después de: ' + ((i.get('c') + ' · ') if i.get('c') else '') + i['n'][:55]
                    opciones.append(etq)
                    self._despues[etq] = i
        self.cb_pos['values'] = opciones
        self.v_pos.set('Al final del grupo')

    def _ok(self):
        n = re.sub(r'\s+', ' ', self.t_n.get('1.0', 'end')).strip()
        c = self.v_c.get().strip()
        if not n:
            self.msg.config(text='Escriba el nombre del insumo.')
            return
        if c and not re.fullmatch(r'[0-9A-Za-z\-]{4,14}', c):
            self.msg.config(text='El código solo lleva números o letras (ej.: 10602050).')
            return
        if self.mes and c and any(i.get('t') == 'i' and i.get('c') == c for i in self.mes['items']):
            if not messagebox.askyesno('Código repetido', f'El código {c} ya está en la matriz de este mes. ¿Agregarlo de todos modos?', parent=self):
                return
        pos = self.v_pos.get()
        self.res = {'c': c, 'n': n, 'nc': self.v_nc.get().strip(), 'um': self.v_um.get().strip(), 'grupo': self.v_g.get().strip(),
                    'inicio': pos == 'Al inicio del grupo', 'despues': clave(self._despues[pos]) if pos in self._despues else None, 'todos': self.v_todos.get()}
        self.destroy()


class ElegirGrupo(simpledialog.Dialog):
    def __init__(self, parent, grupos):
        self.grupos, self.res = grupos, None
        super().__init__(parent, 'Mover a otro grupo')

    def body(self, f):
        ttk.Label(f, text='Grupo de destino:').pack(anchor='w')
        self.v = tk.StringVar(value=self.grupos[0] if self.grupos else '')
        ttk.Combobox(f, textvariable=self.v, values=self.grupos, state='readonly', width=44).pack(pady=4)

    def apply(self):
        self.res = self.v.get()


class DialogoTexto(simpledialog.Dialog):
    def __init__(self, parent, titulo, datos):
        self.datos, self.res = datos, None
        super().__init__(parent, titulo)

    def body(self, f):
        ttk.Label(f, text='Resultado de la revisión:').grid(row=0, column=0, sticky='w')
        self.v_c = tk.StringVar(value=self.datos.get('concl', ''))
        ttk.Combobox(f, textvariable=self.v_c, width=80, values=['Sin señales de alerta en el mes.', 'Se revisaron las señales de alerta y todas quedan justificadas.',
                                                                 'Hay señales de alerta pendientes de justificar.', 'Se encontraron hechos que se informan a la jefatura para su seguimiento.']).grid(row=1, column=0)
        ttk.Label(f, text='Observaciones y justificación:').grid(row=2, column=0, sticky='w')
        self.t_o = tk.Text(f, width=80, height=6)
        self.t_o.insert('1.0', self.datos.get('obs', ''))
        self.t_o.grid(row=3, column=0)
        ttk.Label(f, text='Acciones tomadas o por tomar:').grid(row=4, column=0, sticky='w')
        self.t_a = tk.Text(f, width=80, height=4)
        self.t_a.insert('1.0', self.datos.get('acc', ''))
        self.t_a.grid(row=5, column=0)

    def apply(self):
        self.res = {'concl': self.v_c.get(), 'obs': self.t_o.get('1.0', 'end').strip(), 'acc': self.t_a.get('1.0', 'end').strip()}


def autoprueba(carpeta):
    """Prueba automática (se usa al compilar): crea datos de ejemplo, genera los documentos y sale."""
    ok = False
    try:
        App._inicio = lambda self: None
        app = App(carpeta)
        D = app.D
        items = [{'id': 'g1', 't': 'c', 'n': 'JERINGAS'}, {'id': 'i1', 't': 'i', 'c': '10602050', 'n': 'JERINGA 10 ml PRUEBA', 'ini': 100, 'd': [5, 10]},
                 {'id': 'i2', 't': 'i', 'c': '10602060', 'n': 'JERINGA 20 ml PRUEBA', 'd': [3]}]
        D.crear_mes('PRUEBA', '2026-10', items)
        m = D.mes('PRUEBA', '2026-10')
        m['vales'].append({'id': 'v1', 'num': '1/2026', 'fecha': '2026-10-05', 'rec': '001/2026', 'lineas': [{'item': 'i2', 'cant': 50, 'vto': '2029-01', 'obs': ''}]})
        D.guardar_mes('PRUEBA', '2026-10')
        P = app.pdf
        rutas = [P.matriz('PRUEBA', '2026-10'), P.existencias('PRUEBA', '2026-10'), P.kardex('PRUEBA', '2026-10'), P.hojas('PRUEBA', '2026-10', [1], True),
                 P.recibo('PRUEBA', '2026-10', m['vales'][0]), P.auditoria('PRUEBA', '2026-10', '2026-10'), P.antisoborno('PRUEBA', '2026-10'), P.periodo('PRUEBA', '2026-10', '2026-10')]
        ok = all(os.path.getsize(r) > 1000 for r in rutas)
        app.destroy()
    except Exception as e:
        open(os.path.join(carpeta, 'ERROR.txt'), 'w', encoding='utf-8').write(repr(e))
    open(os.path.join(carpeta, 'RESULTADO.txt'), 'w', encoding='utf-8').write('OK' if ok else 'FALLO')
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == '--autoprueba':
        os.makedirs(sys.argv[2], exist_ok=True)
        autoprueba(sys.argv[2])
    base = sys.argv[1] if len(sys.argv) > 1 and os.path.isdir(sys.argv[1]) else None
    App(base).mainloop()
