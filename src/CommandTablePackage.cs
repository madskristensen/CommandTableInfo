using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CommandTableInfo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(PackageGuids.guidCommandTablePackageString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class CommandTablePackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {   
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await EnableInfoCommand.InitializeAsync(this);
        }
    }
}
