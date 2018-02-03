import { ContextInfo, Helper, List, Types } from "gd-sprest";

/**
 * Team Member Item
 */
export interface ITeamMemberItem extends Types.SP.IListItemQueryResult {
    CheckInStatus: Types.SP.ComplexTypes.FieldManagedMetadataValue;
    CheckInStatus_0: string;
    TeamMember: Types.SP.ComplexTypes.FieldUserValue;
}

/**
 * Datasource
 */
export const Datasource = {
    // Check the cache
    checkCache: () => {
        // See if we have already checked in the user
        let status = sessionStorage.getItem("CheckInDemo");
        return status == "Active";
    },

    // Check the team member in
    checkTeamMemberIn: (item: ITeamMemberItem): PromiseLike<void> => {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field information
            new Helper.ListFormField({
                listName: "Team Members",
                name: "CheckInStatus"
            }).then((fieldInfo: Helper.Types.IListFormMMSFieldInfo) => {
                // Get the term set data
                Helper.ListFormField.loadMMSData(fieldInfo).then(terms => {
                    // Convert the terms into a tree object
                    let termSet = Helper.Taxonomy.toObject(terms);

                    // Get the "Active" status
                    let term = Helper.Taxonomy.findByName(termSet, "Active");
                    if (term) {
                        // Update the status
                        item.update({
                            CheckInStatus: Helper.Taxonomy.toFieldValue(term)
                        }).execute(() => {
                            // Resolve the promise
                            resolve();
                        });
                    }
                });
            });
        });
    },

    // Get the team member
    getTeamMember: (): PromiseLike<ITeamMemberItem> => {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the list
            new List("Team Members")
                // Get the items
                .Items()
                // Set the query:
                // 1) Filter for the current user
                // 2) Include the hidden MMS status field
                // 3) Include the user's full name
                .query({
                    Expand: ["TeamMember"],
                    Filter: "TeamMember eq " + ContextInfo.userId,
                    Select: ["*", "CheckInStatus_0", "TeamMember/Title"]
                })
                // Execute the request
                .execute(items => {
                    // See if the item exists
                    let item = items.results ? items.results[0] as ITeamMemberItem : null;
                    if (item && item.CheckInStatus && item.CheckInStatus_0) {
                        // Update the MMS label
                        // Note - The value returned is the lookup id, not the value.
                        item.CheckInStatus.Label = (item.CheckInStatus_0 || "").split("|")[0];
                    }

                    // Resolve the request
                    resolve(items.results ? items.results[0] : null as any);
                });
        });
    },

    // Update the cache
    updateCache: () => {
        // Set a flag in the session, so we don't run this on every page load
        sessionStorage.setItem("CheckInDemo", "Active");
    }
}