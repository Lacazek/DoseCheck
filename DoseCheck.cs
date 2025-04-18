using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using DoseCheck;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
	public class Script
	{

		public Script()
		{
		}

		[MethodImpl(MethodImplOptions.NoInlining)]
		public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
		{
            UserWindow window = new UserWindow(context.Patient, context.Course, context.PlanSetup, context.PlanSetup.RTPrescription, context.Image);
            try
			{			
				window.ShowDialog();
				window.CloseLog();
			}
			catch 
			{ 
				window.CloseLog(); 
			}
		}	
	}
}
