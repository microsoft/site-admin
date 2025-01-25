import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "site-admin",
    GlobalVariable: "SiteAdmin",
    Lists: {
        Main: "Site Admin Requests"
    },
    ProjectName: "Site Admin",
    ProjectDescription: "Application for adminitrators to make requests for changes on their site collections.",
    SearchFileTypes: "csv doc docx dot dotx pdf pot potx pps ppsx ppt pptx txt xls xlsx xlt xltx",
    SearchTerms: "phi pii ssn",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss",
    Version: "0.1"
};
export default Strings;