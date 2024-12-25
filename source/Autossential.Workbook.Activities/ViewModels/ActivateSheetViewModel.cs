using System.Activities.DesignViewModels;

namespace Autossential.Workbook.Activities.ViewModels
{
    public class ActivateSheetViewModel : DesignPropertiesViewModel
    {
        public ActivateSheetViewModel(IDesignServices services) : base(services)
        {
        }

        public DesignInArgument Sheet { get; set; }

        protected override void InitializeModel()
        {
            base.InitializeModel();
            PersistValuesChangedDuringInit();

            int orderIndex = 0;

            Sheet.IsRequired = true;
            Sheet.IsPrincipal = true;
            Sheet.OrderIndex = orderIndex++;
            Sheet.Placeholder = "Enter the sheet name or zero-based index";
            Sheet.DisplayName = "Sheet Name or Index";
            Sheet.Tooltip = "Enter the sheet name or zero-based index";
        }
    }
}
