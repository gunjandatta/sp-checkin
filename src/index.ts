import "core-js/es6/promise";
import { Configuration } from "./cfg";
import { Datasource } from "./ds";
declare var SP;

/**
 * Check-In Demo
 */
class CheckInDemo {
    // Configuration
    static Configuration = Configuration;

    /**
     * Constructor
     */
    constructor() {
        // Wait for the page to be loaded
        window.addEventListener("load", () => {
            // Wait for the sp.js core script to be loaded, so we can reference the notify and status class
            SP.SOD.executeOrDelayUntilScriptLoaded(() => {
                // Validate the user
                this.validateUser();
            }, "sp.js");
        });
    }

    // Method to validate the user
    private validateUser = () => {
        // Check the cache
        if (Datasource.checkCache()) { return; }

        // Get the user
        Datasource.getTeamMember().then(item => {
            // Ensure the item exists
            if (item) {
                // Ensure the status is active
                let status = (item.CheckInStatus ? item.CheckInStatus.Label : "").toLowerCase();
                if (status != "active") {
                    // Display a status
                    let statusId = SP.UI.Status.addStatus("Checking In", "Welcome " + item.TeamMember.Title + ", we are checking you in.");
                    SP.UI.Status.setStatusPriColor(statusId, "yellow");

                    // Check the user in
                    Datasource.checkTeamMemberIn(item).then(() => {
                        // Clear the statuses
                        SP.UI.Status.removeStatus(statusId);

                        // Display a notification
                        SP.UI.Notify.addNotification("Thank you for checking in.");

                        // Update the session
                        Datasource.updateCache();
                    });
                }
            } else {
                // Display a status
                let statusId = SP.UI.Status.addStatus("Unknown Team Member", "Please contact the site admin to be added to the team member list.");
                SP.UI.Status.setStatusPriColor(statusId, "red");
            }
        });
    }
}

// Make the class available globally
window["CheckInDemo"] = CheckInDemo;

// Create an instance of the class
new CheckInDemo();