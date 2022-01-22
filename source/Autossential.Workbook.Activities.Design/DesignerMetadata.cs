using Autossential.Shared.Activities.Design;
using Autossential.Workbook.Activities.Design.Designers;
using Autossential.Workbook.Activities.Properties;
using System.Activities.Presentation.Metadata;
using System.ComponentModel;

namespace Autossential.Workbook.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public const string MAIN_CATEGORY = "Autossential";
        public const string WORKBOOK = MAIN_CATEGORY + ".Workbook";

        public void Register()
        {
            var workbook = new CategoryAttribute(WORKBOOK);

            ActivitiesAttributesBuilder.Build(Resources.ResourceManager, builder =>
            {
                builder.SetDefaultCategories(
                    Resources.Input_Category,
                    Resources.Output_Category,
                    Resources.InputOutput_Category,
                    Resources.Options_Category);

                builder.Register<GetSheetName, GetSheetNameDesigner>(workbook)
                       .Register<GetSheetNames, GetSheetNamesDesigner>(workbook)
                       .Register<GetHyperlinks, GetHyperlinksDesigner>(workbook)
                       .Register<InsertHyperlink, InsertHyperlinkDesigner>(workbook)
                       .Register<RemoveHyperlinks, RemoveHyperlinksDesigner>(workbook)
                       .Register<WorkbookScope, WorkbookScopeDesigner>(workbook);

                builder.RegisterToMember(
                    new DescriptionAttribute(Resources.ScopeAwareCodeActivity_UseScope),
                    nameof(ScopeAwareCodeActivity<object, object>.UseScope),
                    typeof(ScopeAwareCodeActivity<,>).GetDerivedTypes());
            });
        }
    }
}