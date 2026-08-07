import { DirectorySession, Helper, Types } from "gd-sprest-bs";

// Return type when calling to get group information
export interface IM365Result {
    groups: { [key: string]: Types.SP.Directory.GroupOData };
    error: { [key: string]: boolean };
}

/**
 * M365 Groups
 * Loads and stores information from M365 groups to reduce redundant calls.
 */
export class M365Groups {
    // Group ids that errored when trying to load the information
    private static _errorGroupIds: { [key: string]: boolean } = {};
    static errorLoading(groupId: string): boolean { return this._errorGroupIds[groupId]; }

    // M365 Groups
    private static _groups: { [key: string]: Types.SP.Directory.GroupOData } = {};
    static groupById(groupId: string): Types.SP.Directory.GroupOData { return this._groups[groupId]; }

    // Returns the group id from the login name
    static getGroupId(value: string): string {
        // Get the group id from the login name
        let userInfo = value.split('|');
        let groupInfo = userInfo[userInfo.length - 1];
        let groupId = groupInfo.split('_')[0];

        // Ensure it's a guid and return null if it's not
        return /^[{]?[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}[}]?$/.test(groupId) ? groupId : null;
    }

    // Returns the url to the group
    static getGroupUrl(group: Types.SP.Directory.GroupOData): string {
        return group.calendarUrl.replace("calendar/group", "groups") + "/members";
    }

    // Get the M365 group information
    static getGroupInfo(groupIds: string[], onGroupLoaded?: (group: Types.SP.Directory.GroupOData, groupId: string) => void): PromiseLike<IM365Result> {
        // Return a promise
        return new Promise((resolve) => {
            let result: IM365Result = { groups: {}, error: {} };

            // Parse the group ids
            Helper.Executor(groupIds, (groupId) => {
                // See if we already loaded this group and it errored
                if (this._errorGroupIds[groupId]) {
                    result.error[groupId] = true;
                    return;
                }

                // See if we already loaded this group
                if (this._groups[groupId]) {
                    result.groups[groupId] = this._groups[groupId];
                    return;
                }

                // Return a promise
                return new Promise(resolve => {
                    // Load the group information
                    this.loadGroupById(groupId).then(group => {
                        // Save the group information
                        result.groups[groupId] = group;

                        // Call the event
                        onGroupLoaded ? onGroupLoaded(group, groupId) : null;

                        // Try the next group
                        resolve(null);
                    }, () => {
                        // Add the error
                        result.error[groupId] = true;

                        // Call the event
                        onGroupLoaded ? onGroupLoaded(null, groupId) : null;

                        // Try the next group
                        resolve(null);
                    });
                });
            }).then(() => { resolve(result); });
        });
    }

    // Loads the m365 group
    private static loadGroupById(groupId: string): PromiseLike<Types.SP.Directory.GroupOData> {
        let group: Types.SP.Directory.GroupOData = { id: groupId } as Types.SP.Directory.GroupOData;

        // Return a promise
        return new Promise((resolve) => {
            // Get the group information
            let getGroupInfo = (): PromiseLike<void> => {
                return new Promise(resolve => {
                    DirectorySession().group(groupId).query({
                        Select: ["calendarUrl", "displayName", "id", "isPublic", "mail"]
                    }).execute(group => {
                        // Update the group information
                        group.calendarUrl = group.calendarUrl;
                        group.displayName = group.displayName;
                        group.id = group.id;
                        group.isPublic = group.isPublic;
                        group.mail = group.mail;

                        // Resolve the group information
                        resolve();
                    }, () => {
                        // Error loading this group
                        this._errorGroupIds[groupId] = true;

                        // Resolve the group information
                        resolve();
                    });
                });
            }

            // Get the group members/owner information
            let getGroupUsers = (member: boolean): PromiseLike<void> => {
                return new Promise(resolve => {
                    DirectorySession().group(groupId)[member ? "members" : "owners"]().query({
                        GetAllItems: true,
                        Top: 5000,
                        Select: ["principalName", "id", "displayName", "mail"]
                    }).execute(users => {
                        // Set the users
                        group[member ? "members" : "owners"] = users as any;

                        // Resolve the users
                        resolve();
                    }, () => {
                        // Resolve the request
                        resolve();
                    }, true);
                });
            }

            // Get the group information, members and users
            Promise.all([getGroupInfo(), getGroupUsers(true), getGroupUsers(false)]).then(() => {
                // Set the group
                this._groups[group.id] = group;

                // Resolve the request
                resolve(this._groups[group.id]);
            });
        });
    }
}