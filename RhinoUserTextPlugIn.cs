using System.Collections.Generic;
using Rhino;
using Rhino.PlugIns;
using Rhino.UI;

namespace RhinoUserText
{
    ///<summary>
    /// <para>Every RhinoCommon .rhp assembly must have one and only one PlugIn-derived
    /// class. DO NOT create instances of this class yourself. It is the
    /// responsibility of Rhino to create an instance of this class.</para>
    /// <para>To complete plug-in information, please also see all PlugInDescription
    /// attributes in AssemblyInfo.cs (you might need to click "Project" ->
    /// "Show All Files" to see it in the "Solution Explorer" window).</para>
    ///</summary>
    public class RhinoUserTextPlugIn : Rhino.PlugIns.PlugIn

    {

        public override PlugInLoadTime LoadTime { get; }

        public RhinoUserTextPlugIn()
        {
            LoadTime = PlugInLoadTime.AtStartup;
            Instance = this;
          
        }
        
        ///<summary>Gets the only instance of the RhinoUserTextPlugIn plug-in.</summary>
        public static RhinoUserTextPlugIn Instance
        {
            get; private set;
        }

        protected override void DocumentPropertiesDialogPages(RhinoDoc doc, List<OptionsDialogPage> pages)
        {
            var page = new Views.UserStringsDocumentOptionsPage(doc);
            pages.Add(page);
        }

        protected override void ObjectPropertiesPages(List<ObjectPropertiesPage> pages)
        {
            var page = new Views.UserStringsObjectPropertiesPage();
            pages.Add(page);
        }


        // You can override methods here to change the plug-in behavior on
        // loading and shut down, add options pages to the Rhino _Option command
        // and mantain plug-in wide options in a document.
    }
}