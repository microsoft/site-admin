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
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                EnableAttachments: false,
                ReadSecurity: SPTypes.ListReadSecurity.User,
                WriteSecurity: SPTypes.ListWriteSecurity.User
            },
            TitleFieldDisplayName: "Site Url",
            TitleFieldDescription: "Enter the relative/absolute url to the site collection you want to enable custom scripts on.",
            CustomFields: [
                {
                    name: "ProcessFlag",
                    title: "Process Flag",
                    type: Helper.SPCfgFieldType.Boolean,
                    defaultValue: "0",
                    showInDisplayForm: false,
                    showInEditForm: false,
                    showInListSettings: false,
                    showInNewForm: false,
                    showInViewForms: false
                },
                {
                    name: "RequestType",
                    title: "Request Type",
                    type: Helper.SPCfgFieldType.Choice,
                    defaultValue: "",
                    required: true,
                    showInEditForm: false,
                    choices: [
                        "App Catalog", "Client Side Assets", "Custom Script", "Company Wide Sharing Links",
                        "Increase Storage", "Lock State", "No Crawl", "Custom Search Property", "Restrict Content Discovery",
                        "Teams Connected"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "RequestValue",
                    title: "Request Value",
                    type: Helper.SPCfgFieldType.Text
                },
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
                        "New", "Cancelled", "Error", "Processed", "Completed"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewQuery: '<OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy>',
                    ViewFields: [
                        "ID", "Author", "RequestType", "LinkTitle", "RequestValue", "Created", "Status"
                    ]
                }
            ]
        }
    ]
});