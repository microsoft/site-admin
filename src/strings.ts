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
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.1"
};
export default Strings;