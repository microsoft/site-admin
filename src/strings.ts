import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, cloudEnv?:string, sourceUrl?: string) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the cloud enivronment and source url
    Strings.CloudEnvironment = cloudEnv || Strings.CloudEnvironment;
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
}

/**
 * Global Constants
 */
const Strings = {
    AppElementId: "site-admin",
    CloudEnvironment: "",
    GlobalVariable: "SiteAdmin",
    Lists: {
        Main: "Site Admin Requests"
    },
    ProjectName: "Site Admin",
    ProjectDescription: "Application for adminitrators to make requests for changes on their site collections.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss",
    Version: "0.1"
};
export default Strings;