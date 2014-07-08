using System.Diagnostics;
using System.Management.Automation;

namespace Microsoft.Samples.PowerShell.Commands
{ // Windows PowerShell namespace.

    #region GetProcCommand

    /// <summary>
    /// This class implements the Get-Proc cmdlet.
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "Proc", DefaultParameterSetName = "Id")]
    public class GetProcCommand : Cmdlet
    {
        [Parameter(ParameterSetName = "Id")]
        [Parameter(ParameterSetName = "OneOnly")]
        public int Id { get; set; }

        [Parameter(ParameterSetName = "OneOnly")]
        public SwitchParameter OneOnly { get; set; }

        [Parameter]
        public SwitchParameter DoNothing { get; set; }

        #region Cmdlet Overrides

        /// <summary>
        /// The ProcessRecord method calls the Process.GetProcesses 
        /// method to retrieve the processes of the local computer. 
        /// Then, the WriteObject method writes the associated processes 
        /// to the pipeline.
        /// </summary>
        protected override void ProcessRecord()
        {
            // Retrieve the current processes.
            var processes = Process.GetProcesses();

            // Write the processes to the pipeline to make them available
            // to the next cmdlet. The second argument (true) tells Windows 
            // PowerShell to enumerate the array and to send one process 
            // object at a time to the pipeline.
            WriteObject(processes, true);
        }

        #endregion Overrides
    } // End GetProcCommand class.

    #endregion GetProcCommand
}