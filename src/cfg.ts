import { Helper, SPTypes } from "gd-sprest";

/**
 * Configuration
 */
export const Configuration = {
    // List
    List: new Helper.SPConfig({
        ListCfg: [
            {
                ListInformation: {
                    BaseTemplate: SPTypes.ListTemplateType.GenericList,
                    Description: "Sample list for the check-in demo.",
                    Title: "Team Members"
                },
                TitleFieldDisplayName: "Role",
                CustomFields: [
                    // Team Member
                    {
                        name: "TeamMember",
                        title: "Team Member",
                        type: Helper.SPCfgFieldType.User,
                        selectionMode: SPTypes.FieldUserSelectionType.PeopleOnly
                    } as Helper.Types.IFieldInfoUser,
                    // Status
                    {
                        name: "CheckInStatus",
                        title: "Check-In Status",
                        type: Helper.SPCfgFieldType.MMS
                    }
                ],
                ViewInformation: [
                    // Default View
                    {
                        ViewName: "All Items",
                        ViewFields: ["TeamMember", "LinkTitle", "CheckInStatus"]
                    }
                ]
            }
        ]
    }),

    // Web Custom Action
    Web: new Helper.SPConfig({
        CustomActionCfg: {
            Web: [
                {
                    Description: "Enables the automated check-in demo.",
                    Location: "ScriptLink",
                    Name: "CheckInDemo",
                    ScriptSrc: "~site/siteassets/checkin/check-in.js",
                    Title: "Check-In Demo"
                }
            ]
        }
    })
}