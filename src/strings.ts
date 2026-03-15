import { ContextInfo, SPTypes } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, maxBatchSize: number, maxRequests: number) => {
    // Set the context
    ContextInfo.setPageContext(context.pageContext);

    // Update the source url
    Strings.MaxBatchSize = typeof (maxBatchSize) === "number" ? maxBatchSize : Strings.MaxBatchSize;
    Strings.MaxRequests = typeof (maxRequests) === "number" ? maxRequests : Strings.MaxRequests;
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
    MaxBatchSize: 25,
    MaxRequests: 1,
    ProjectName: "Site Admin",
    ProjectDescription: "Application for adminitrators to make requests for changes on their site collections.",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TimeFormat: "YYYY-MMM-DD HH:mm:ss",
    Version: "0.9.7"
};
export default Strings;