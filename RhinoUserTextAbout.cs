using System;
using Rhino;
using Rhino.Commands;

namespace RhinoUserText
{
    public class RhinoUserTextAbout : Command
    {
        static RhinoUserTextAbout _instance;
        public RhinoUserTextAbout()
        {
            _instance = this;
        }

        ///<summary>The only instance of the RhinoUserTextAbout command.</summary>
        public static RhinoUserTextAbout Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "RhinoUserTextAbout"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.
            return Result.Success;
        }
    }
}
