import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { App } from "./app";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { InstallationModal } from "./install";
import { Security } from "./security";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// App Properties
interface IAppProps {
    context?: any;
    el: HTMLElement;
    disableProps?: string[];
}

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
            //DataSource.init(props.azureFunctionUrl)
            DataSource.init(null)
        ]).then(
            // Success
            () => {
                // Create the application
                new App(props.el, props.disableProps);
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
    GlobalVariable.render({ el: elApp });
}