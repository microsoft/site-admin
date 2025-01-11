import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App, IAppProps } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    appDescription: Strings.ProjectDescription,
    render: (props: IAppProps) => {
        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context);

            // Update the configuration
            Configuration.setWebUrl(ContextInfo.webServerRelativeUrl);
        }

        // Initialize the application and load the theme
        Promise.all([
            ThemeManager.load(true),
            DataSource.init()
        ]).then(
            // Success
            () => {
                // Clear the element
                while (props.el.firstChild) { props.el.removeChild(props.el.firstChild); }

                // Create the application
                new App(props);
            },

            // Error
            () => {
                // See if the user has the correct permissions
                Security.hasPermissions().then(hasPermissions => {
                    // See if the user has permissions
                    if (hasPermissions) {
                        // Show the installation modal
                        InstallationModal.show();
                    }
                });
            }
        );
    },
    updateTheme: (themeInfo) => { ThemeManager.update(themeInfo); }
};

// Make is available in the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render({ el: elApp, siteProps: {}, webProps: {} });
}