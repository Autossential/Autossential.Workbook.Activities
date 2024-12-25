using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class GetActiveSheetViewModel : DesignPropertiesViewModel
    {
        public GetActiveSheetViewModel(IDesignServices services) : base(services)
        {
        }

        public DesignOutArgument<int> SheetIndex { get; set; }
        public DesignOutArgument<string> SheetName { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            SheetIndex.IsRequired = false;
            SheetIndex.IsPrincipal = false;
            SheetIndex.OrderIndex = orderIndex++;
            SheetIndex.DisplayName = "Index";
            SheetIndex.Tooltip = "The zero-based index of the active sheet";

            SheetName.IsRequired = false;
            SheetName.IsPrincipal = false;
            SheetName.OrderIndex = orderIndex++;
            SheetName.DisplayName = "Name";
            SheetIndex.Tooltip = "The name of the active sheet";
        }
    }
}
