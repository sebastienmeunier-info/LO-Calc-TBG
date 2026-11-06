# PoC LibreOffice Calc : tableaux de bord (DataPilot + graphiques + segment)
import uno
import unohelper

from com.sun.star.frame import XDispatchProvider, XDispatch
from com.sun.star.awt import Rectangle, MessageBoxButtons
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.PushButtonType import OK, CANCEL
from com.sun.star.sheet.DataPilotFieldOrientation import ROW, COLUMN, DATA, PAGE
from com.sun.star.sheet.GeneralFunction import SUM
from com.sun.star.table import CellAddress, CellRangeAddress

IMPL_NAME = "com.example.lo-dashboard-ext.Handler"
CMD_PREFIX = "com.example.lo-dashboard-ext:"
DASH_SHEET = "Dashboard"
DASH_COLUMN_STEP = 12  # colonnes entre chaque bloc pivot


def _ctx():
    import uno
    lc = uno.getComponentContext()
    return lc


def _desktop(ctx):
    return ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)


def _get_calc_doc(ctx):
    doc = _desktop(ctx).CurrentComponent
    if doc is None or not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return None
    return doc


def _msgbox(doc, text, title="Tableau de bord"):
    if doc is None:
        return
    try:
        frame = doc.CurrentController.Frame
        win = frame.ContainerWindow
        box = win.getToolkit().createMessageBox(
            win, MESSAGEBOX, MessageBoxButtons.BUTTONS_OK, title, text)
        box.execute()
    except Exception:
        pass


def _used_range(sheet):
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    ra = cursor.RangeAddress
    return ra


def _column_names(sheet, used):
    names = []
    for c in range(used.StartColumn, used.EndColumn + 1):
        names.append(sheet.getCellByPosition(c, used.StartRow).getString())
    return names


def _classify_columns(sheet, used):
    text_cols, num_cols = [], []
    for c in range(used.StartColumn, used.EndColumn + 1):
        numeric = False
        for r in range(used.StartRow + 1, min(used.EndRow, used.StartRow + 30) + 1):
            cell = sheet.getCellByPosition(c, r)
            if cell.getType().value in (1, 3):  # VALUE ou FORMULA
                numeric = True
                break
        (num_cols if numeric else text_cols).append(c)
    return text_cols, num_cols


def _ensure_sheet(doc, name):
    if doc.Sheets.hasByName(name):
        return doc.Sheets.getByName(name)
    doc.Sheets.insertNewByName(name, doc.Sheets.Count)
    return doc.Sheets.getByName(name)


def _insert_pivot(doc, dptables, name, target_col, source_range,
                  row_field, data_field, col_field=None, page_field=None):
    desc = dptables.createDataPilotDescriptor()
    desc.setSourceRange(source_range)
    fields = desc.getDataPilotFields()
    fields.getByIndex(row_field).Orientation = ROW
    fields.getByIndex(data_field).Orientation = DATA
    fields.getByIndex(data_field).Function = SUM
    if col_field is not None:
        fields.getByIndex(col_field).Orientation = COLUMN
    if page_field is not None:
        fields.getByIndex(page_field).Orientation = PAGE
    target = CellAddress()
    target.Sheet = doc.Sheets.getByName(DASH_SHEET).RangeAddress.Sheet
    target.Column = target_col
    target.Row = 0
    dptables.insertNewByName(name, target, desc)
    table = dptables.getByName(name)
    return table.getOutputRange()


def _add_chart(dash_sheet, name, output_range, x, y):
    charts = dash_sheet.Charts
    rect = Rectangle(x, y, 9000, 6000)
    charts.addNewByName(name, rect, (output_range,), True, True)


def _clear_dashboard(doc, dash_sheet):
    for name in list(dash_sheet.DataPilotTables.ElementNames):
        dash_sheet.DataPilotTables.removeByName(name)
    for name in list(dash_sheet.Charts.ElementNames):
        dash_sheet.Charts.removeByName(name)
    dash_sheet.clearContents(
        1 | 2 | 4 | 8 | 16 | 32 | 64 | 128 | 256 | 512 | 1024)


def create_dashboard(ctx):
    doc = _get_calc_doc(ctx)
    if doc is None:
        _msgbox(None, "Ouvrez un classeur Calc contenant une feuille de données.")
        return
    if doc.Sheets.Count == 0:
        return
    data_sheet = doc.Sheets.getByIndex(0)
    used = _used_range(data_sheet)
    if used.EndRow < 1 or used.EndColumn < 1:
        _msgbox(doc, "La première feuille ne contient pas assez de données "
                     "(1 ligne d'en-têtes + des données).")
        return

    text_cols, num_cols = _classify_columns(data_sheet, used)
    if len(text_cols) < 1 or len(num_cols) < 1:
        _msgbox(doc, "Il faut au moins une colonne de catégories (texte) "
                     "et une colonne numérique.")
        return

    source = CellRangeAddress()
    source.Sheet = data_sheet.RangeAddress.Sheet
    source.StartColumn = used.StartColumn
    source.StartRow = used.StartRow
    source.EndColumn = used.EndColumn
    source.EndRow = used.EndRow

    dash_sheet = _ensure_sheet(doc, DASH_SHEET)
    _clear_dashboard(doc, dash_sheet)
    dptables = dash_sheet.DataPilotTables

    # Pivot 1 : catégories (colonne texte 1) en lignes, somme numérique 1
    out1 = _insert_pivot(doc, dptables, "TDB_Pivot1", 0, source,
                         text_cols[0], num_cols[0])
    _add_chart(dash_sheet, "TDB_Chart1", out1, 100, 100)

    # Pivot 2 : 2e dimension en lignes (ou en colonnes si disponible)
    if len(text_cols) >= 2:
        out2 = _insert_pivot(doc, dptables, "TDB_Pivot2", DASH_COLUMN_STEP,
                             source, text_cols[1], num_cols[0])
    else:
        out2 = _insert_pivot(doc, dptables, "TDB_Pivot2", DASH_COLUMN_STEP,
                             source, text_cols[0], num_cols[0],
                             col_field=num_cols[0] if len(num_cols) > 1 else None)
    _add_chart(dash_sheet, "TDB_Chart2", out2, 100, 7000)

    # Pivot 3 : 1re dimension en lignes, 2e numérique en données
    out3 = _insert_pivot(doc, dptables, "TDB_Pivot3", 2 * DASH_COLUMN_STEP,
                         source, text_cols[0],
                         num_cols[1] if len(num_cols) > 1 else num_cols[0])
    _add_chart(dash_sheet, "TDB_Chart3", out3, 7000, 100)

    _msgbox(doc, "Tableau de bord créé dans la feuille « %s » : "
                 "3 tables croisés dynamiques + 3 graphiques." % DASH_SHEET)


def _unique_values(sheet, used, col):
    values = []
    seen = set()
    for r in range(used.StartRow + 1, used.EndRow + 1):
        v = sheet.getCellByPosition(col, r).getString()
        if v not in seen:
            seen.add(v)
            values.append(v)
    return values


def _pick_value_dialog(ctx, title, label, values):
    smgr = ctx.ServiceManager
    toolkit = smgr.createInstanceWithContext(
        "com.sun.star.awt.Toolkit", ctx)
    dialog = smgr.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog.Width = 180
    dialog.Height = 160
    dialog.Title = title

    lbl = dialog.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    lbl.PositionX = 8
    lbl.PositionY = 6
    lbl.Width = 160
    lbl.Height = 12
    lbl.Label = label
    dialog.insertByName("lbl", lbl)

    lb_model = dialog.createInstance(
        "com.sun.star.awt.UnoControlListBoxModel")
    lb_model.PositionX = 8
    lb_model.PositionY = 22
    lb_model.Width = 160
    lb_model.Height = 100
    lb_model.StringItemList = tuple(values)
    lb_model.MultipleSelection = False
    lb_model.SelectedItems = (0,)
    dialog.insertByName("lb", lb_model)

    ok_btn = dialog.createInstance(
        "com.sun.star.awt.UnoControlButtonModel")
    ok_btn.PositionX = 40
    ok_btn.PositionY = 128
    ok_btn.Width = 50
    ok_btn.Height = 16
    ok_btn.PushButtonType = OK
    ok_btn.Label = "Appliquer"
    dialog.insertByName("ok", ok_btn)

    cancel_btn = dialog.createInstance(
        "com.sun.star.awt.UnoControlButtonModel")
    cancel_btn.PositionX = 95
    cancel_btn.PositionY = 128
    cancel_btn.Width = 50
    cancel_btn.Height = 16
    cancel_btn.PushButtonType = CANCEL
    cancel_btn.Label = "Annuler"
    dialog.insertByName("cancel", cancel_btn)

    ctrl = smgr.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialog", ctx)
    ctrl.setModel(dialog)
    ctrl.createPeer(toolkit, None)
    if ctrl.execute() != 1:  # OK
        return None
    sel = ctrl.getControl("lb").getSelectedItems()
    if not sel:
        return None
    return values[sel[0]]


def apply_slicer(ctx):
    doc = _get_calc_doc(ctx)
    if doc is None:
        _msgbox(None, "Ouvrez un classeur Calc.")
        return
    data_sheet = doc.Sheets.getByIndex(0)
    used = _used_range(data_sheet)
    names = _column_names(data_sheet, used)

    col_name = _pick_value_dialog(
        ctx, "Segment - champ", "Champ à filtrer :", names)
    if col_name is None:
        return
    col = names.index(col_name)
    values = ["(Tous)"] + _unique_values(data_sheet, used, col)
    value = _pick_value_dialog(
        ctx, "Segment - valeur", "Valeur de « %s » :" % col_name, values)
    if value is None:
        return

    applied = 0
    for s in doc.Sheets:
        dptables = s.DataPilotTables
        for name in dptables.ElementNames:
            table = dptables.getByName(name)
            desc_fields = table.DataPilotDescriptor.getDataPilotFields()
            for i in range(desc_fields.Count):
                field = desc_fields.getByIndex(i)
                if field.Name != col_name:
                    continue
                try:
                    if value == "(Tous)":
                        field.setPropertyValue("SelectedPage", "")
                    else:
                        field.setPropertyValue("SelectedPage", value)
                    table.refresh()
                    applied += 1
                except Exception:
                    pass
    _msgbox(doc, "Segment appliqué sur %d table(s) croisée(s) dynamique(s)."
                 % applied)


def refresh_tables(ctx):
    doc = _get_calc_doc(ctx)
    if doc is None:
        return
    n = 0
    for s in doc.Sheets:
        dptables = s.DataPilotTables
        for name in dptables.ElementNames:
            dptables.getByName(name).refresh()
            n += 1
    _msgbox(doc, "%d table(s) actualisée(s)." % n)


class DashboardHandler(unohelper.Base, XDispatchProvider, XDispatch):
    def __init__(self, ctx):
        self.ctx = ctx

    # XDispatchProvider
    def queryDispatch(self, url, target, flags):
        if url.Protocol == CMD_PREFIX:
            return self
        return None

    def queryDispatches(self, requests):
        out = []
        for r in requests:
            d = self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
            out.append(d)
        return tuple(out)

    # XDispatch
    def dispatch(self, url, args):
        cmd = url.Path
        try:
            if cmd == "create":
                create_dashboard(self.ctx)
            elif cmd == "slicer":
                apply_slicer(self.ctx)
            elif cmd == "refresh":
                refresh_tables(self.ctx)
        except Exception as e:
            _msgbox(_get_calc_doc(self.ctx), "Erreur : %s" % e)

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass


g_ImplementationHelper = unohelper.ImplementationContainer()
g_ImplementationHelper.addImplementation(
    DashboardHandler,
    IMPL_NAME,
    ("com.sun.star.frame.ProtocolHandler",),
)
