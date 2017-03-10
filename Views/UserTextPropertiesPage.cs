using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Eto.Drawing;
using Eto.Forms;
using Rhino;
using Rhino.DocObjects;
using Rhino.Input;
using Rhino.Input.Custom;
using Rhino.UI;


namespace RhinoUserText.Views
{
    #region Object User Strings Properties Page
    class UserStringsObjectPropertiesPage : ObjectPropertiesPage
    {
        private UserStringsPanelControl m_page_control;
        public override string EnglishPageTitle => LOC.STR("Object Strings");
        public override System.Drawing.Icon Icon => Properties.Resources.Notes;
        public override object PageControl => (m_page_control ?? (m_page_control = new UserStringsPanelControl()));
        public override bool ShouldDisplay(RhinoObject rhObj)
        {
            return true;
        }
        public override void InitializeControls(RhinoObject rhObj)
        {
            m_page_control?.InitializeControls(SelectedObjects);
        }
        public override bool OnActivate(bool active)
        {
          
            m_page_control.LoadObjectStrings(SelectedObjects);
            return base.OnActivate(active);

        }
    }
    #endregion

    #region Document User Strings Options Dialog
    public class UserStringsDocumentOptionsPage : OptionsDialogPage
    {
        public RhinoDoc Document { get; set; }
        public UserStringsDocumentOptionsPage(RhinoDoc doc) : base(LOC.STR("User Strings"))
        {
            Document = doc;
        }
        private UserStringsPanelControl m_page_control;
        public override object PageControl => (m_page_control ?? (m_page_control = new UserStringsPanelControl(Document)));
    }
    #endregion

    class UserStringsPanelControl : Panel
    {
        #region localized strings for things
        private readonly string m_dialog_import_title = LOC.STR("Import CSV");
        private readonly string m_dialog_export_title = LOC.STR("Export CSV");
        private readonly string m_csvstr = LOC.STR("CSV (Comma delimited");
        private readonly string m_txtstr = LOC.STR("TXT (Comma delimited)");
        private readonly string m_file_error_message_title = LOC.STR("Import Error");
        private readonly string m_file_locked_message = LOC.STR("File is currently in use.");
        #endregion

        private GridView m_grid;
        private Dictionary<string, string> m_dictionary = new Dictionary<string, string>();
        private Collection<UserStringItem> m_collection = new ObservableCollection<UserStringItem>();
        private RhinoObject[] SelectedObjects { get; set; }
        private bool IsDocumentText { get;}
        private RhinoDoc Document { get;}


        //Panel Creation
        #region Panel Creation
        public UserStringsPanelControl(RhinoDoc doc)
        {   //Called when creating the Document User Strings Panel
            Document = doc;
            IsDocumentText = true;
            LayoutPanel();
            LoadDocumentStrings();
        }
        public UserStringsPanelControl()
        {   //Called when Creating a panel for Object User Strings
            IsDocumentText = false;
            LayoutPanel();
        }
        public void InitializeControls(RhinoObject[] rhObjs)
        {
           

            SelectedObjects = rhObjs;
            LoadObjectStrings(SelectedObjects);
        }
        #endregion

        //Panel Layout
        #region Panel Layout
        private void LayoutPanel()
        {
            #region Construct a GridView
            string headerkey = LOC.STR("Key");
            string headervalue = LOC.STR("Value");

            m_grid = new GridView { DataStore = m_collection };

            m_grid.SizeChanged += Grid_SizeChanged;
            m_grid.GridLines = GridLines.Both;
            m_grid.AllowColumnReordering = false;
            m_grid.AllowMultipleSelection = true;
            m_grid.Columns.Add(new GridColumn
            {
                DataCell = new TextBoxCell("Key"),
                HeaderText = headerkey,
                AutoSize = false,
            });
            m_grid.Columns.Add(new GridColumn
            {
                DataCell = new TextBoxCell("Value"),
                HeaderText = headervalue,
                AutoSize = false,
            });

            m_grid.Columns[0].Editable = true;
            m_grid.Columns[1].Editable = true;
            m_grid.Columns[0].Sortable = true;
            m_grid.Columns[1].Sortable = true;

            if (IsDocumentText)
            {
                m_grid.CellEdited += Grid_Doc_CellEdited;
            }
            else
            {
                m_grid.CellEdited += Grid_Obj_CellEdited;
            }
            #endregion


            var btn_stack_layout = CreateButtonStackLayout();

            //Build the main table
            Content = new TableLayout
            {
                Spacing = new Size(4, 4),
                Rows =
                {
                   // new TableRow(new LabelSeparator() {Text = LOC.STR("User Text")}),
                    new TableLayout()
                    {
                        Spacing = new Size(10,4),
                        Rows =
                        {
                        new TableRow(new TableCell(m_grid,true), new TableCell(btn_stack_layout)){ ScaleHeight = true},
                        }
                    },
                }
            };
          
        }
        public StackLayout CreateButtonStackLayout()
        {

            #region Create Buttons for the Object Properties Panel
            //Add our document buttons
            var items = new Collection<StackLayoutItem>();
            var btn_obj_add = new Button { Text = LOC.STR("New...") };
            btn_obj_add.Click += Btn_obj_add_Click;
            var btn_obj_delete = new Button { Text = LOC.STR("Delete...") };
            btn_obj_delete.Click += Btn_obj_delete_Click;
            var btn_obj_export = new Button { Text = LOC.STR("Export...") };
            btn_obj_export.Click += Btn_obj_export_Click;
            var btn_obj_import = new Button { Text = LOC.STR("Import...") };
            btn_obj_import.Click += Btn_obj_import_Click;
            var btn_obj_match = new Button { Text = LOC.STR("Match...") };
            btn_obj_match.Click += Btn_obj_match_Click;
            #endregion

            #region Create Buttons for the Panel
            //Add our document buttons
            var btn_doc_add = new Button { Text = LOC.STR("New...") };
            btn_doc_add.Click += BtnAdd_Click;
            var btn_doc_delete = new Button { Text = LOC.STR("Delete...") };
            btn_doc_delete.Click += BtnDelete_Click;
            var btn_doc_export = new Button { Text = LOC.STR("Export...") };
            btn_doc_export.Click += BtnExport_Click;
            var btn_doc_import = new Button { Text = LOC.STR("Import...") };
            btn_doc_import.Click += Btn_import_Click;
            #endregion
            
            #region Content Layout

            //Create our stack layout based on which type of panel we need (doc strings or object strings)
            StackLayout stack_layout;
            if (IsDocumentText)
            {
                stack_layout = new StackLayout
                {
                    Padding = 4,
                    HorizontalContentAlignment = HorizontalAlignment.Right,
                    Spacing = 4,
                    Items =
                    {
                        btn_doc_add,
                        btn_doc_delete,
                        btn_doc_import,
                        btn_doc_export,
                    }

                };
            }
            else
            {
                stack_layout = new StackLayout
                {
                    Padding = 4,
                    HorizontalContentAlignment = HorizontalAlignment.Right,
                    Spacing = 4,
                    Items =
                    {
                        btn_obj_add,
                        btn_obj_delete,
                        btn_obj_import,
                        btn_obj_export,
                        btn_obj_match,
                    }

                };
            }
            #endregion

            return stack_layout;
        }
        #endregion
        
        //All Panel Buttons
        #region All Panel Buttons
        //Object Properties Panel Buttons
        #region Object Properties Panel Buttons
        private void Btn_obj_add_Click(object sender, EventArgs e)
        {
            m_collection.Add(new UserStringItem { Key = string.Empty, Value = string.Empty });
            m_grid.BeginEdit(m_collection.Count - 1, 0);
        }
        private void Btn_obj_delete_Click(object sender, EventArgs e)
        {

        }
        private void Btn_obj_import_Click(object sender, EventArgs e)
        {
           ImportCsv(ref m_collection);
            foreach(var ro in SelectedObjects)
            { 
                foreach (var c in m_collection)
                {
                    ro.Attributes.SetUserString(c.Key, c.Value);
                    ro.CommitChanges();
                }
            }
        }
        private void Btn_obj_export_Click(object sender, EventArgs e)
        {
          ExportCsv(m_collection);
        }
        private void Btn_obj_match_Click(object sender, EventArgs e)
        {

            var go = new GetObject();
            go.Get();
            if (go.Result() != GetResult.Object)
                return;

            var dict = new Dictionary<string, string>();
            GetUserStringDictionary(go.Object(0).Object(), ref dict);

            foreach (var ro in SelectedObjects)
            {
                //TODO steve to update attributes
            }


        }
        #endregion
        
        //Document Properties Panel Buttons
        #region Document Properties Panel Buttons
        private void Btn_import_Click(object sender, System.EventArgs e)
        {
            //Import a csv file to Document Strings 
            ImportCsv(ref m_collection);
            foreach (var c in m_collection)
                Document.Strings.SetString(c.Key, c.Value);
        }

        private void BtnExport_Click(object sender, System.EventArgs e)
        {
            //Export the Document Strings to a CSV
           ExportCsv(m_collection);
        }
        private void BtnDelete_Click(object sender, System.EventArgs e)
        {
            //DELETE DOCUMENT STRINGS
            while (m_grid.SelectedRows.Any())
            {
                var first_selected = m_grid.SelectedRows.First();
                if (first_selected < 0)
                    break;

                var dcount = Document.Strings.Count;
                for (var i = 0; i <= dcount; i++)
                {
                    var key = Document.Strings.GetKey(i);
                    var value = Document.Strings.GetValue(i);
                    if (key != m_collection[first_selected].Key || value != m_collection[first_selected].Value)
                        continue;

                    Document.Strings.Delete(key);
                    break;
                }
                m_collection.RemoveAt(first_selected);

            }


        }
        private void BtnAdd_Click(object sender, System.EventArgs e)
        {
            //Add a new Document String
            m_collection.Add(new UserStringItem { Key = string.Empty, Value = string.Empty });
            m_grid.BeginEdit(m_collection.Count - 1, 0);

        }
        #endregion
        #endregion

        //Grid Events
        #region Grid Events
        private void Grid_Obj_CellEdited(object sender, GridViewCellEventArgs e)
        {
            //a grid item was just edited for an Object Strings
            var rc = e.Row;
            var unique = CreateUniqueKey(m_collection[rc].Key, false,ref m_collection);
            m_collection[rc].Key = unique;
            foreach (var ro in SelectedObjects)
            {
                ro.Attributes.SetUserString(m_collection[rc].Key, m_collection[rc].Value);
                ro.CommitChanges();
            }
        }

        private void Grid_Doc_CellEdited(object sender, GridViewCellEventArgs e)
        {
            //A grid item for a Document String was edited
            var rc = m_collection[e.Row];
            if (e.Column == 0)
            {
                if (rc.Key == string.Empty) return;
                var key = CreateUniqueKey(rc.Key, false, ref m_collection);
                //Update the documents strings with newly entered string
                m_collection[e.Row].Key = $"{key}";
            }
            Document.Strings.SetString(rc.Key, rc.Value);
           
        }


        private void Grid_SizeChanged(object sender, System.EventArgs e)
        {
            //RESIZE OUR COLUMN WIDTHS TO FIT THE GRID
            if (m_grid.Columns.Count != 2) return;
            var first_or_default = m_grid.Columns.FirstOrDefault();
            if (first_or_default != null)
                first_or_default.Width = Convert.ToInt16(m_grid.Width / 4);
            if (first_or_default != null) m_grid.Columns.Last().Width = Convert.ToInt16((m_grid.Width - (m_grid.Width / 4)) - 2);
        }

        #endregion

        //Private Methods
        #region Private Methods
        public void LoadObjectStrings(RhinoObject[] selectedObjects)
        {
            m_dictionary = new Dictionary<string, string>();
            SelectedObjects = selectedObjects;
            if (!selectedObjects.Any()) return;

            m_dictionary = new Dictionary<string, string>();

            foreach (var ro in selectedObjects)
                  GetUserStringDictionary(ro,ref m_dictionary);

            m_collection.Clear();
            foreach (var entry in m_dictionary)
                m_collection.Add(new UserStringItem {Key = entry.Key, Value = entry.Value});

            m_grid.DataStore = m_collection;


        }

        public void LoadDocumentStrings()
        {
            m_collection.Clear();
            for (int i = 0; i < Document.Strings.Count; i++)
            {
                m_collection.Add(new UserStringItem {Key = Document.Strings.GetKey(i),Value = Document.Strings.GetValue(i)});
            }
        }
        private void GetUserStringDictionary(RhinoObject ro, ref Dictionary<string, string> dictionary)
        {
            var sc = ro.Attributes.GetUserStrings().Count;
            for (var i = 0; i < sc; i++)
            {
                var k = ro.Attributes.GetUserStrings().GetKey(i);
                var v = ro.Attributes.GetUserString(k);
                if (k == string.Empty) continue;

                if (dictionary.ContainsKey(k))
                {
                    string value;
                    m_dictionary.TryGetValue(k, out value);
                    if (value != v)
                        m_dictionary[k] = "(varies)";
                }
                else
                {
                    dictionary.Add(k, v);
                }
            }
        }
        private void ImportCsv(ref Collection<UserStringItem> collection)
        {
            //Import a csv file 

            try
            {
                #region ETO Open File Dialog
                var fd = new Eto.Forms.OpenFileDialog
                {
                    MultiSelect = false,
                    Title = m_dialog_import_title
                };

                fd.Filters.Add(new FileFilter(m_csvstr, ".csv"));
                fd.Filters.Add(new FileFilter(m_txtstr, ".txt"));
                fd.CurrentFilter = fd.Filters[0];
                var result = fd.ShowDialog(RhinoEtoApp.MainWindow);
                if (result != DialogResult.Ok) return;

                var file_name = fd.FileName;

                if (IsFileLocked(new FileInfo(file_name)))
                {
                    Dialogs.ShowMessage(m_file_locked_message, m_file_error_message_title,
                        ShowMessageButton.OK, ShowMessageIcon.Error);
                    return;
                }
                #endregion

                #region Stream Reader to parse opened file
                using (var reader = new StreamReader(file_name))
                {
                    string line;
                    var csv_parser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

                    while ((line = reader.ReadLine()) != null)
                    {
                        var x = csv_parser.Split(line);
                        if (!x.Any()) continue;

                        var entry = new UserStringItem { Key = CreateUniqueKey(x[0], true, ref collection), Value = x[1] };
                        collection.Add(entry);
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                Rhino.Runtime.HostUtils.DebugString("Exception caught during Options - User Text - CSV Import");
                Rhino.Runtime.HostUtils.ExceptionReport(ex);
            }
        }
        private void ExportCsv(Collection<UserStringItem> collection)
        {

            //Save off a set of keys
            try
            {
                var fd = new Eto.Forms.SaveFileDialog();
                fd.Filters.Add(new FileFilter(m_csvstr, ".csv"));
                fd.Filters.Add(new FileFilter(m_txtstr, ".txt"));
                fd.Title = m_dialog_export_title;
                var export_result = fd.ShowDialog(RhinoEtoApp.MainWindow);

                if (export_result != DialogResult.Ok)
                    return;

                if (fd.CheckFileExists)
                {
                    var replace = Dialogs.ShowMessage(LOC.STR($"Are you sure you want to replace {fd.FileName}?"),
                        m_dialog_export_title, ShowMessageButton.YesNoCancel, ShowMessageIcon.Question);

                    if (replace != ShowMessageResult.Yes)
                        return;

                    if (IsFileLocked(new FileInfo(fd.FileName)))
                    {
                        Dialogs.ShowMessage(m_file_locked_message, m_file_error_message_title,
                                ShowMessageButton.OK, ShowMessageIcon.Error);
                        return;
                    }

                    File.Delete(fd.FileName);

                }


                var csv = new StringBuilder();
                foreach (var entry in collection.Select(co => $"{co.Key},{co.Value}"))
                    csv.AppendLine(entry);

                File.WriteAllText(fd.FileName, csv.ToString());
            }
            catch (Exception ex)
            {
                Rhino.Runtime.HostUtils.DebugString("Exception caught during Options - User Text - CSV File Export");
                Rhino.Runtime.HostUtils.ExceptionReport(ex);
            }



        }
        private string CreateUniqueKey(string key, bool importMode, ref Collection<UserStringItem> collection)
        {
           // key = key.Replace(@"\", "\\");

            var duplicate_count = collection.Count(c => c.Key == key);

            if (importMode)
            {
                if (duplicate_count == 0)
                    return key;
            }
            else
            {
                if (duplicate_count <= 1)
                    return key;
            }

            var i = 1;
            var test_key = $"{key} ({i})";
            while (i < 100000)
            {
                test_key = $"{key} ({i})";
                var count = collection.Count(c => c.Key == test_key);
                if (count == 0)
                    break;
                i++;
            }
            return test_key;
        }
        private  bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
      
        #endregion
    }


    public class UserStringItem
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }
}
