import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { ISiteGroupInfo, IWeb } from "@pnp/sp/presets/all";
import { IGroupCreateFormValues } from "../components/GroupCreatePanel/GroupCreatePanel";
import * as strings from "GroupDashboardWebPartStrings";
import { IGroupEditFormValues } from "../components/GroupEditPanel/GroupEditPanel";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

/**
 * Get all web's groups
 */
export async function getGroups(web: IWeb = null): Promise<ISiteGroupInfo[]> {
	if (!web) {
		web = sp.web;
	}
	return web.siteGroups.select("Id", "Title", "Description").get();
}

export async function createGroup(data: IGroupCreateFormValues): Promise<ISiteGroupInfo> {
	const addedGroup = await sp.web.siteGroups.add({
		...data
	});
	return addedGroup.group.select("Id", "Title", "Description").get();
}

export async function deleteGroup(id: number): Promise<void> {
	const usersCount = (await sp.web.siteGroups.getById(id).users.get()).length;
	if (usersCount > 0) {
		throw new Error(strings.DeleteGroupErrorUsersExists.replace("${id}", id.toString()));
	}
	return sp.web.siteGroups.removeById(id);
}

export async function updateGroup(id: number, data: Partial<IGroupEditFormValues>): Promise<ISiteGroupInfo> {
	const editedGroup = await sp.web.siteGroups.getById(id).update({
		...data,
	});
	return editedGroup.group.select("Id", "Title", "Description").get();
}

export async function getGroupUsers(id: number): Promise<ISiteUserInfo[]> {
	return sp.web.siteGroups.getById(id).users.select("Id", "Title", "Email", "LoginName").get();
}

export async function getUsers(): Promise<ISiteUserInfo[]> {
	return sp.web.siteUsers.select("Id", "Title", "Email", "LoginName").usingCaching().get();
}

export async function addUserToGroup(userLoginName: string, groupId: number): Promise<ISiteUserInfo> {
	return (await sp.web.siteGroups.getById(groupId).users.add(userLoginName)).select("Id", "Title", "Email", "LoginName").get();
}

export async function deleteUserFromGroup(userEmail: string, groupId: number): Promise<void> {
	return sp.web.siteGroups.getById(groupId).users.getByEmail(userEmail).delete();
}