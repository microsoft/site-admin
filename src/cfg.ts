import { Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    ListCfg: [
        {
            ListInformation: {
                Title: Strings.Lists.Main,
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            TitleFieldDisplayName: "Site Url",
            TitleFieldDescription: "Enter the relative/absolute url to the site collection you want to enable custom scripts on.",
            CustomFields: [
                {
                    name: "Status",
                    title: "Status",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "New",
                    indexed: true,
                    required: true,
                    showInNewForm: false,
                    showInEditForm: false,
                    choices: [
                        "New", "Error", "Processed", "Completed"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: [
                        "ID", "LinkTitle", "Created", "Status"
                    ]
                }
            ]
        }
    ]
});