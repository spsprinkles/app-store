import { ListSecurity, ListSecurityDefaultGroups } from "dattatable";
import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * Security
 * Code related to the security groups the user belongs to.
 */
export class Security {
    private static _listSecurity: ListSecurity = null;

    // Admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }
    private static _adminGroup: Types.SP.Group = null;
    static get AdminGroup(): Types.SP.Group { return this._adminGroup; }

    // Developers
    private static _isDeveloper: boolean = false;
    static get IsDeveloper(): boolean { return this._isDeveloper; }
    private static _developerGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: false,
        AutoAcceptRequestToJoinLeave: true,
        Description: Strings.SecurityGroups.Managers.Description,
        Title: Strings.SecurityGroups.Managers.Name,
        OnlyAllowMembersViewMembership: false
    };
    private static _developerGroup: Types.SP.Group = null;
    static get DeveloperGroup(): Types.SP.Group { return this._developerGroup; }

    // Manager
    private static _isManager: boolean = false;
    static get IsManager(): boolean { return this._isManager; }
    private static _managerGroupInfo: Types.SP.GroupCreationInformation = {
        AllowMembersEditMembership: false,
        Description: Strings.SecurityGroups.Managers.Description,
        Title: Strings.SecurityGroups.Managers.Name,
        OnlyAllowMembersViewMembership: false
    };
    private static _managerGroup: Types.SP.Group = null;
    static get ManagerGroup(): Types.SP.Group { return this._managerGroup; }

    // SecurityGroupUrl
    private static _securityGroupUrl = ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=";
    static get SecurityGroupUrl(): string { return this._securityGroupUrl };

    // Initialization
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Create the list security
            this._listSecurity = new ListSecurity({
                groups: [
                    this._developerGroupInfo, this._managerGroupInfo
                ],
                listItems: [
                    {
                        listName: Strings.Lists.Main,
                        groupName: ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Main,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    {
                        listName: Strings.Lists.Main,
                        groupName: this._managerGroupInfo.Title,
                        permission: SPTypes.RoleType.Contributor
                    },
                    {
                        listName: Strings.Lists.Ratings,
                        groupName: ListSecurityDefaultGroups.Owners,
                        permission: SPTypes.RoleType.Administrator
                    },
                    {
                        listName: Strings.Lists.Ratings,
                        groupName: ListSecurityDefaultGroups.Visitors,
                        permission: SPTypes.RoleType.Reader
                    },
                    {
                        listName: Strings.Lists.Ratings,
                        groupName: this._managerGroupInfo.Title,
                        permission: SPTypes.RoleType.Contributor
                    },
                ],
                onGroupsLoaded: () => {
                    // Set the groups
                    this._adminGroup = this._listSecurity.getGroup(ListSecurityDefaultGroups.Owners);
                    this._managerGroup = this._listSecurity.getGroup(this._managerGroupInfo.Title);

                    // See if the user belongs to the group
                    this._isAdmin = this._listSecurity.CurrentUser.IsSiteAdmin || this._listSecurity.isInGroup(ContextInfo.userId, ListSecurityDefaultGroups.Owners);
                    this._isManager = this._listSecurity.isInGroup(ContextInfo.userId, this._managerGroupInfo.Title);

                    // Ensure all of the groups exist
                    if (this._adminGroup && this._managerGroup) {
                        // Get the owner
                        this._managerGroup.Owner().execute(owner => {
                            // See if they are the admin group
                            if (owner.Id != this._adminGroup.Id) {
                                // Log
                                console.log(`[${this.ManagerGroup.Title} Group] The owner is not correct. Updating it now.`);

                                // Set the group owner
                                Helper.setGroupOwner(this._managerGroup.Title, this._adminGroup.Title).then(() => {
                                    // Log
                                    console.log(`[${this._managerGroup.Title} + " Group] The owner was updated successfully to ${this._adminGroup.Title}.`);

                                    // Resolve the request
                                    resolve();
                                }, resolve);
                            } else {
                                // Resolve the request
                                resolve();
                            }
                        }, reject);
                    } else {
                        // Reject the request
                        reject();
                    }
                }
            })
        });
    }

    // Displays the security group configuration
    static show(onComplete: () => void) {
        // Create the groups
        this._listSecurity.show(true, onComplete);
    }
}