using UiPath.Studio.Activities.Api;

namespace Autossential.Workbook.Activities.API
{
    public class DesignerMetadata : IRegisterWorkflowDesignApi
    {
        public void Initialize(IWorkflowDesignApi api)
        {
            api.ScopedActivitiesService.SetScopedActivities(typeof(WorkbookScope), [
                typeof(ReadRange),
                typeof(WriteRange),
                typeof(ReadCell),
                typeof(WriteCell)
            ]);
        }
    }
}
