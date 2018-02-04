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
        // Wait for the notification class to be made available, so we can reference the status class
        SP.SOD.executeOrDelayUntilScriptLoaded(() => {
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
        }, "sp.js");
    }
}

// Make the class available globally
window["CheckInDemo"] = CheckInDemo;

// Create an instance of the class
new CheckInDemo();