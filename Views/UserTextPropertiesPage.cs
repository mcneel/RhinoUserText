using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
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
using System.Runtime.InteropServices;
using Rhino.UI.Controls;

namespace RhinoUserText.Views
{


    #region Object User Strings Properties Page
    class UserStringsObjectPropertiesPage : ObjectPropertiesPage
    {
        private UserStringsPanelControl m_page_control;
        public override string EnglishPageTitle => LOC.STR("Attribute Text");
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
        public UserStringsDocumentOptionsPage(RhinoDoc doc) : base(LOC.STR("Document Text"))
        {
            Document = doc;
        }
        private UserStringsPanelControl m_page_control;
        public override object PageControl => (m_page_control ?? (m_page_control = new UserStringsPanelControl(Document)));
        public override bool OnApply()
        {
            return (m_page_control == null || m_page_control.OnApply());
        }
    }
    #endregion


    [Guid("A9B10690-89F5-4AAC-9445-DF90D56F5E2D")]
    public class UserStringsPanelControl : Panel, IPanel
    {
        //This Panel is used in 3 places
        //Rhino Options - Document Text
        //Object Properties - Attribute Text
        //Floating - Document Text (like the notes panel)
        //Depending on which panel mode its in the controls and events will shift around. 


        #region localized strings for things
        private readonly string m_new = LOC.STR("New...");
        private readonly string m_delete = LOC.STR("Delete..");
        private readonly string m_import = LOC.STR("Import...");
        private readonly string m_export = LOC.STR("Export...");
        private readonly string m_match = LOC.STR("Match...");
        private readonly string m_command_prompt = LOC.STR("Select object To match");
        private readonly string m_headerkey = LOC.STR("Key");
        private readonly string m_headervalue = LOC.STR("Value");
        private readonly string m_dialog_import_title = LOC.STR("Import CSV");
        private readonly string m_dialog_export_title = LOC.STR("Export CSV");
        private readonly string m_csvstr = LOC.STR("CSV (Comma delimited");
        private readonly string m_txtstr = LOC.STR("TXT (Comma delimited)");
        private readonly string m_file_error_message_title = LOC.STR("Import Error");
        private readonly string m_file_locked_message = LOC.STR("File is currently in use.");
        private readonly string m_replace_file = LOC.STR("Are you sure you want to replace");
        private readonly string m_file_success = LOC.STR("File successfully written to");
        private readonly string m_varies = LOC.STR("(varies)");
        private readonly string m_csv = LOC.STR(".csv");
        private readonly string m_txt = LOC.STR(".txt");
        private readonly string m_attribute_txt = LOC.STR("Attribute Text");
        private readonly string m_document_txt = LOC.STR("Document Text");
        private readonly string m_filter_txt = LOC.STR("Filter");
        #endregion

        private GridView m_grid;
        private Dictionary<string, string> m_dictionary = new Dictionary<string, string>();
        private Collection<UserStringItem> m_collection = new ObservableCollection<UserStringItem>();
        private RhinoObject[] SelectedObjects { get; set; }
        private RhinoDoc Document { get; set; }
        private bool IsDocumentText { get;}
        private bool UpdateStrings { get; set; }
        private bool DocumentPanelActive { get; set; }
        public static Guid PanelId => typeof(UserStringsPanelControl).GUID;
      
        //Panel Creation
        #region Panel Creation
        public UserStringsPanelControl(RhinoDoc doc)
        {   //Called when creating the Document Text Panel in Document Options
            Document = doc;
            IsDocumentText = true;
            LayoutPanel();
            LoadDocumentStrings();
        }


        public void InitializeControls(RhinoObject[] rhObjs)
        {
            //Called for Attribute Panel Mode
            SelectedObjects = rhObjs;
            LoadObjectStrings(SelectedObjects);
        }
        public UserStringsPanelControl()
        {   //Called when Creating a panel for Object Attribute Text
            IsDocumentText = false;
            LayoutPanel();
        }

       

       
        #endregion

        //Panel Layout
        #region Panel Layout
        private void LayoutPanel()
        {
            #region Construct a GridView
            m_grid = new GridView { DataStore = m_collection };
            m_grid.SizeChanged += Grid_SizeChanged;
            m_grid.GridLines = GridLines.Both;
            m_grid.AllowColumnReordering = false;
            m_grid.AllowMultipleSelection = true;
           
            m_grid.Columns.Add(new GridColumn
            {
                DataCell = new TextBoxCell("Key"),
                HeaderText = m_headerkey,
                AutoSize = false,
                
            });
            m_grid.Columns.Add(new GridColumn
            {
                DataCell = new TextBoxCell("Value"),
                HeaderText = m_headervalue,
                AutoSize = false,
            });

            m_grid.Columns[0].Editable = true;
            m_grid.Columns[1].Editable = true;
            m_grid.Columns[0].Sortable = true;
            m_grid.Columns[1].Sortable = true;
            m_grid.Height = 200;

            if (IsDocumentText)
            {
                m_grid.CellEdited += Grid_Doc_CellEdited;
            }
            else
            {
                m_grid.CellEdited += Grid_Obj_CellEdited;
            }

            if (DocumentPanelActive)
                m_grid.CellEdited += Grid_Doc_CellEdited;

            #endregion

            var btn_stack_layout = CreateButtonStackLayout();
            //Build the main table

            var separator_txt = IsDocumentText ? m_document_txt : m_attribute_txt;

            //Filter button
            var btn_filter = new Button { Text = m_filter_txt};
            btn_filter.Click += Btn_filter_Click;

            var txt_filter = new TextBox();
            txt_filter.KeyUp += Txt_filter_KeyUp;

            Content =
                new TableLayout()
                {
                    Padding = 10,
                    Spacing = new Size(4,4),
                    Rows =
                    {
                        new LabelSeparator() {Text = separator_txt},
                        new StackLayout
                        {
                            Spacing = 4,
                            Orientation = Orientation.Horizontal,
                            HorizontalContentAlignment = HorizontalAlignment.Right,
                            Items =
                            {
                                txt_filter,
                                btn_filter,
                            }
                        },
                        new TableLayout()
                          {
                            Padding = new Padding(4, 4, 4, 20),
                            Spacing = new Size(4, 4),
                            Rows =
                            {
                                new TableLayout()
                                {
                                    Spacing = new Size(4, 4),
                                    Rows =
                                    {
                                        new TableRow(new TableCell(m_grid, true), new TableCell(btn_stack_layout))
                                        {
                                            ScaleHeight = false
                                        }
                                    }
                                }
                            }
                         }
                    }
                };
        }

        private void Txt_filter_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Keys.Enter)
                return;

            RhinoApp.WriteLine("TODO : Enter was pressed");
        }

        public StackLayout CreateButtonStackLayout()
        {

            #region Create Buttons for the Object Properties Panel
            //Add our document buttons
            var btn_obj_add = new Button { Text = m_new};
            btn_obj_add.Click += Btn_obj_add_Click;
            var btn_obj_delete = new Button { Text = m_delete};
            btn_obj_delete.Click += Btn_obj_delete_Click;
            var btn_obj_export = new Button { Text = m_export};
            btn_obj_export.Click += Btn_obj_export_Click;
            var btn_obj_import = new Button { Text = m_import};
            btn_obj_import.Click += Btn_obj_import_Click;
            var btn_obj_match = new Button { Text = m_match};
            btn_obj_match.Click += Btn_obj_match_Click;
            #endregion

            #region Create Buttons for the Panel
            //Add our document buttons
            var btn_doc_add = new Button { Text = m_new};
            btn_doc_add.Click += Btn_doc_add_Click;
            var btn_doc_delete = new Button { Text = m_delete};
            btn_doc_delete.Click += Btn_doc_delete_Click;
            var btn_doc_export = new Button { Text = m_export};
            btn_doc_export.Click += Btn_doc_export_Click;
            var btn_doc_import = new Button { Text = m_import};
            btn_doc_import.Click += Btn_doc_import_Click;
            #endregion
            
            #region Content Layout

            //Create our stack layout based on which type of panel we need (doc strings or object strings)
            StackLayout stack_layout;
            if (IsDocumentText)
            {
                //Options - Document Properties Panel
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
                //Object Properties Panel
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
        
        //Filter Grid Button
        private void Btn_filter_Click(object sender, EventArgs e)
        {
            RhinoApp.WriteLine("Filter it");
        }
            
            
        //Object Properties Panel Buttons
        #region Object Properties Panel Buttons
        private void Btn_obj_add_Click(object sender, EventArgs e)
        {
            m_collection.Add(new UserStringItem { Key = string.Empty, Value = string.Empty });
            m_grid.BeginEdit(m_collection.Count - 1, 0);
        }
        private void Btn_obj_delete_Click(object sender, EventArgs e)
        {
            //Delete a string entry from any object thats selected
            while (m_grid.SelectedRows.Any())
            {
                var first_selected = m_grid.SelectedRows.First();
                if (first_selected < 0)
                    break;
                    foreach (var ro in SelectedObjects)
                    {
                        ro.Attributes.DeleteUserString(m_collection[first_selected].Key);
                        ro.CommitChanges();
                    }
                m_collection.RemoveAt(first_selected);
             }
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
            var originalselected_objects = SelectedObjects;
            var first_or_default = originalselected_objects.FirstOrDefault();
            if (first_or_default == null) return;

            var doc = first_or_default.Document;
            
            doc.Objects.UnselectAll(false);
            doc.Views.Redraw();

            var go = new GetObject();
            go.SetCommandPrompt(m_command_prompt);
            go.Get();
           
            if (go.Result() != GetResult.Object)
                return;

            var dict = new Dictionary<string, string>();
            GetUserStringDictionary(go.Object(0).Object(), ref dict);

            doc.Objects.UnselectAll(false);
            foreach (var ro in originalselected_objects)
            {
                doc.Objects.Select(ro.Id);
            }
            doc.Views.Redraw();

            foreach (var ro in SelectedObjects)
            {
               ro.Attributes.DeleteAllUserStrings();
                foreach (var entry in dict)
                {
                    ro.Attributes.SetUserString(entry.Key, entry.Value);
                }
                ro.CommitChanges();
            }
        }
        #endregion

        //Document Properties Panel Buttons
        #region Document Properties Panel Buttons
        private void Btn_doc_add_Click(object sender, EventArgs e)
        {
            //Add a new Document String
            m_collection.Add(new UserStringItem { Key = string.Empty, Value = string.Empty });
            m_grid.BeginEdit(m_collection.Count - 1, 0);

        }
        private void Btn_doc_delete_Click(object sender, EventArgs e)
        {
            //DELETE COLLECTION STRINGS
            while (m_grid.SelectedRows.Any())
            {
                var first_selected = m_grid.SelectedRows.First();
                if (first_selected < 0)
                    break;

                if (DocumentPanelActive)
                    Document.Strings.Delete(m_collection[first_selected].Key);

                m_collection.RemoveAt(first_selected);
            }
        }
        private void Btn_doc_import_Click(object sender, EventArgs e)
        {
            //Import a csv file to Document Strings 
            ImportCsv(ref m_collection);
            foreach (var c in m_collection)
                Document.Strings.SetString(c.Key, c.Value);
        }
        private void Btn_doc_export_Click(object sender, EventArgs e)
        {
            //Export the Document Strings to a CSV
           ExportCsv(m_collection);
        }


        //OnApply Button
        public bool OnApply()
        {
            //Remove all existing strings
            while (Document.Strings.Count > 0)
            {
                Document.Strings.Delete(Document.Strings.GetKey(0));
            }


            //Save all new strings
            foreach (var entry in m_collection)
            {
             //   Debug.WriteLine($"{entry.Key} / {entry.Value}");
                Document.Strings.SetString(entry.Key, entry.Value);
            }

            Debug.WriteLine("Apply Button - Delete Doc Strings and Reload Collection from Doc");
            return true;
        }
        #endregion

       
        #endregion

        //Grid Events
        #region Grid Events
        private void Grid_Obj_CellEdited(object sender, GridViewCellEventArgs e)
        {
            //a grid item was just edited for an Objects User Strings
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
                m_collection[e.Row].Key = key;
            }

                if (!DocumentPanelActive) return;

            var entry = m_collection[e.Row];
            Document.Strings.SetString(entry.Key, entry.Value);
            Debug.WriteLine("GRID CELL EDITED");
            Debug.WriteLine("================");
        }
        private void Grid_SizeChanged(object sender, EventArgs e)
        {
            //RESIZE OUR COLUMN WIDTHS TO FIT THE GRID
            if (m_grid.Columns.Count != 2) return;
            var first_or_default = m_grid.Columns.FirstOrDefault();
            if (first_or_default != null)
                first_or_default.Width = Convert.ToInt16(m_grid.Width / 4);
            if (first_or_default != null) m_grid.Columns.Last().Width = Convert.ToInt16((m_grid.Width - (m_grid.Width / 4)) - 2);
        }

        #endregion
        
        //Helpers
        #region Helpers

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
            //Read all of the document strings into our main collection and grid controls
            if (Document == null) return;
            m_collection.Clear();
            for (var i = 0; i < Document?.Strings.Count; i++)
            {
                m_collection.Add(new UserStringItem {Key = Document?.Strings.GetKey(i),Value = Document?.Strings.GetValue(i)});
            }
            UpdateStrings = false;
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
                        m_dictionary[k] = m_varies;
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

                fd.Filters.Add(new FileFilter(m_csvstr, m_csv));
                fd.Filters.Add(new FileFilter(m_txtstr, m_txt));
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
        private void ExportCsv(IEnumerable<UserStringItem> collection)
        {

            //Save off a set of keys
            try
            {
                var fd = new Eto.Forms.SaveFileDialog();
                fd.Filters.Add(new FileFilter(m_csvstr, m_csv));
                fd.Filters.Add(new FileFilter(m_txtstr, m_txt));
                fd.Title = m_dialog_export_title;
                var export_result = fd.ShowDialog(RhinoEtoApp.MainWindow);

                if (export_result != DialogResult.Ok)
                    return;

                if (fd.CheckFileExists)
                {
                    var replace = Dialogs.ShowMessage($"{m_replace_file} {fd.FileName}",
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
                RhinoApp.WriteLine($"{m_file_success} {fd.FileName}");
            }
            catch (Exception ex)
            {
                Rhino.Runtime.HostUtils.DebugString("Exception caught during Options - User Text - CSV File Export");
                Rhino.Runtime.HostUtils.ExceptionReport(ex);
            }



        }

        private static string CreateUniqueKey(string key, bool importMode, ref Collection<UserStringItem> collection)
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
        private static bool IsFileLocked(FileInfo file)
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
                stream?.Close();
            }

            //file is not locked
            return false;
        }

        #endregion
        
        //Floating Document Text Panel Mode
        #region Floating Document Text Panel
        public static UserStringsPanelControl Panel(uint documentSerialNumber)
        {
            return (Panels.GetPanel(PanelId, documentSerialNumber) as UserStringsPanelControl);
        }
        //Floating Document Text Panel Constructor
        public UserStringsPanelControl(uint documentSerialNumber)
        {
            //Called when creating a panel for the Document Text ("like notes")
            Document = RhinoDoc.FromRuntimeSerialNumber(documentSerialNumber);
            DocumentPanelActive = true;
            IsDocumentText = true;
            LayoutPanel();
            LoadDocumentStrings();
        }
        //Floating Document Text Panel Events 
        #region document text IPanel events
        public void PanelShown(uint documentSerialNumber, ShowPanelReason reason)
        {
            Document = RhinoDoc.FromRuntimeSerialNumber(documentSerialNumber);
            RhinoDoc.DocumentPropertiesChanged += RhinoDoc_DocumentPropertiesChanged;
            RhinoApp.Idle += RhinoApp_Idle;
            DocumentPanelActive = true;
            LayoutPanel();
        }
        public void PanelHidden(uint documentSerialNumber, ShowPanelReason reason)
        {
            DocumentPanelActive = false;
            RhinoDoc.DocumentPropertiesChanged -= RhinoDoc_DocumentPropertiesChanged;
            RhinoApp.Idle -= RhinoApp_Idle;
        }

        public void PanelClosing(uint documentSerialNumber, bool onCloseDocument)
        {
            DocumentPanelActive = false;
            RhinoDoc.DocumentPropertiesChanged -= RhinoDoc_DocumentPropertiesChanged;
            RhinoApp.Idle -= RhinoApp_Idle;
        }

        private void RhinoApp_Idle(object sender, EventArgs e)
        {
            //if the document properties have changed let's go look at the doc strings.
            if (UpdateStrings)
                LoadDocumentStrings();
        }

        private void RhinoDoc_DocumentPropertiesChanged(object sender, DocumentEventArgs e)
        {
            //The document has been changed, we need to look at our doc strings the next time rhino is idle.
            UpdateStrings = true;
        }

        #endregion
       
        #endregion 



    }


    public class UserStringItem
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }
}
