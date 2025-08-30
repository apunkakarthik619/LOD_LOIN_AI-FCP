# RPS_CreateAndBind_LOIN_AI_Params_Adaptive_TxSafe.py
# Creates shared parameters (if missing) and binds them at the INSTANCE level.
#   Will create: LOIN_Stage (Text), LOIN_Pass (Yes/No), LOIN_Missing (Text),
#   Also creates: LOD_Score (Number) and AI_FinalStatus (Text).
# Uses standard BuiltInParameterGroup to be compatible with multiple Revit versions. If a group isn’t available, falls back to a safe default.
# Transaction-safe: only starts a new transaction if the doc isn't already modifiable.

import os
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
app   = __revit__.Application

# ---------------- CONFIG ----------------
LOIN_PASS_AS_YESNO = True  # False → make LOIN_Pass as Text
BIC_LIST = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_DuctCurves,
    DB.BuiltInCategory.OST_PipeCurves,
    DB.BuiltInCategory.OST_CableTray
]
SP_GROUP_NAME = "LOIN_AI"  # Shared-parameters group name in your SP file
# ---------------------------------------

# --------- Version-adaptive group token ----------
USE_GROUPTYPEID = False
PARAM_GROUP_TOKEN = None
try:
    PARAM_GROUP_TOKEN = DB.BuiltInParameterGroup.PG_DATA
except:
    try:
        PARAM_GROUP_TOKEN = DB.GroupTypeId.Data
        USE_GROUPTYPEID = True
    except:
        UI.TaskDialog.Show("Create/Bind Params",
                           "Neither BuiltInParameterGroup nor GroupTypeId available.\n"
                           "Please share your Revit version/API details.")
        raise SystemExit

# --------- Helpers ----------
def make_category_set(doc, bic_list):
    cats = DB.CategorySet()
    for bic in bic_list:
        cat = doc.Settings.Categories.get_Item(bic)
        if cat:
            cats.Insert(cat)
    return cats

def ensure_shared_params_file(app):
    sp_path = app.SharedParametersFilename
    if not sp_path or not os.path.exists(sp_path):
        UI.TaskDialog.Show(
            "Create/Bind Params",
            "Set a valid Shared Parameters file:\nManage → Shared Parameters → Browse…"
        )
        raise SystemExit
    sp_defs = app.OpenSharedParameterFile()
    if sp_defs is None:
        UI.TaskDialog.Show(
            "Create/Bind Params",
            "Could not open the Shared Parameters file. Ensure it is accessible and try again."
        )
        raise SystemExit
    return sp_defs

def get_or_create_group(sp_defs, group_name):
    for g in sp_defs.Groups:
        if g.Name == group_name:
            return g
    return sp_defs.Groups.Create(group_name)

def get_spec_id_for_text():
    return DB.SpecTypeId.String.Text

def get_spec_id_for_yesno():
    return DB.SpecTypeId.Boolean.YesNo

def get_spec_id_for_number():
    return DB.SpecTypeId.Number

def find_definition_by_name(sp_defs, name):
    for g in sp_defs.Groups:
        for d in g.Definitions:
            if d.Name == name:
                return d
    return None

def create_external_definition(group, name, spec_id):
    opts = DB.ExternalDefinitionCreationOptions(name, spec_id)
    opts.Visible = True
    return group.Definitions.Create(opts)

def ensure_definitions(sp_defs, group):
    defs = {}
    spec_text   = get_spec_id_for_text()
    spec_yesno  = get_spec_id_for_yesno()
    spec_number = get_spec_id_for_number()

    targets = [
        ("LOIN_Stage",     spec_text),
        ("LOIN_Pass",      spec_yesno if LOIN_PASS_AS_YESNO else spec_text),
        ("LOIN_Missing",   spec_text),
        ("LOD_Score",      spec_number),
        ("AI_FinalStatus", spec_text),
    ]
    for name, spec in targets:
        d = find_definition_by_name(sp_defs, name)
        if d is None:
            d = create_external_definition(group, name, spec)
        defs[name] = d
    return defs

def insert_or_reinsert_binding(param_bindings, definition, binding, group_token):
    """
    Handles both overloads:
      - Insert(Definition, Binding, BuiltInParameterGroup)
      - Insert(Definition, Binding, GroupTypeId)
    """
    try:
        if not param_bindings.Insert(definition, binding, group_token):
            param_bindings.ReInsert(definition, binding, group_token)
        return True
    except:
        try:
            alt = DB.GroupTypeId.Data if not USE_GROUPTYPEID else DB.BuiltInParameterGroup.PG_DATA
            if not param_bindings.Insert(definition, binding, alt):
                param_bindings.ReInsert(definition, binding, alt)
            return True
        except Exception as ex2:
            UI.TaskDialog.Show("Create/Bind Params",
                               "Failed to bind parameter '{}':\n{}".format(definition.Name, ex2))
            return False

def instance_bind_all(doc, defs, catset, group_token):
    pb = doc.ParameterBindings
    ib = DB.InstanceBinding(catset)
    for _, d in defs.items():
        insert_or_reinsert_binding(pb, d, ib, group_token)

def catset_is_empty(catset):
    try:
        return catset.IsEmpty
    except:
        it = catset.ForwardIterator(); it.Reset()
        return not it.MoveNext()

# --------- MAIN ---------
# Read-only guard
if getattr(doc, "IsReadOnly", False):
    UI.TaskDialog.Show("Create/Bind Params",
                       "The document is read-only. Save a writable copy or obtain write access, then re-run.")
    raise SystemExit

sp_defs = ensure_shared_params_file(app)
group   = get_or_create_group(sp_defs, SP_GROUP_NAME)
defs    = ensure_definitions(sp_defs, group)
catset  = make_category_set(doc, BIC_LIST)

if catset_is_empty(catset):
    UI.TaskDialog.Show("Create/Bind Params",
                       "No valid categories to bind. Check BIC_LIST in the script.")
    raise SystemExit

# Transaction-safe block
started_here = False
tx = DB.Transaction(doc, "Create & Bind LOIN/AI parameters")
try:
    if not doc.IsModifiable:
        tx.Start()
        started_here = True
    # Do the binding work
    instance_bind_all(doc, defs, catset, PARAM_GROUP_TOKEN)
    if started_here:
        tx.Commit()
    else:
        # We were inside an existing transaction; nothing to commit here.
        pass
except Exception as ex:
    if started_here:
        tx.RollBack()
    UI.TaskDialog.Show("Create/Bind Params", "Failed during binding:\n{}".format(ex))
    raise

UI.TaskDialog.Show(
    "Create/Bind Params",
    "Done.\nCreated/bound parameters:\n"
    "- LOIN_Stage (Text)\n"
    "- LOIN_Pass ({})\n"
    "- LOIN_Missing (Text)\n"
    "- LOD_Score (Number)\n"
    "- AI_FinalStatus (Text)".format("Yes/No" if LOIN_PASS_AS_YESNO else "Text")
)
