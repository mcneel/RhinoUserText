using Rhino;
using Rhino.Commands;
using Rhino.UI;

namespace RhinoUserText.Commands
{
    public class DocumentText : Command
    {
        static DocumentText _instance;

        public DocumentText()
        {
            Panels.RegisterPanel(PlugIn,typeof(Views.UserStringsPanelControl),Rhino.UI.LOC.STR("Document Text"),Properties.Resources.Notes);
            _instance = this;
        }

        ///<summary>The only instance of the DocumentText command.</summary>
        public static DocumentText Instance => _instance;

        public override string EnglishName => "DocumentText";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            var panel_id = Views.UserStringsPanelControl.PanelId;
            var visible = Panels.IsPanelVisible(panel_id);

            if (mode == RunMode.Scripted)
            {
                var go = new Rhino.Input.Custom.GetOption();
                go.SetCommandPrompt(LOC.STR("Choose document text option"));
                var hide_index = go.AddOption("Hide");
                var show_index = go.AddOption("Show");
                var toggle_index = go.AddOption("Toggle");
                go.Get();
                if (go.CommandResult() !=Result.Success)
                    return go.CommandResult();

                var option = go.Option();
                if (null == option)
                    return Result.Failure;

                var index = option.Index;
                if (index == hide_index)
                {
                    if (visible)
                        Panels.ClosePanel(panel_id);
                }
                else if (index == show_index)
                {
                    if (!visible)
                        Panels.OpenPanel(panel_id);
                }
                else if (index == toggle_index)
                {
                    if (visible)
                        Panels.ClosePanel(panel_id);
                    else
                        Panels.OpenPanel(panel_id);
                }

            }
            else
            {
                Panels.OpenPanel(panel_id);
            }



            return Result.Success;
        }
    }
}
