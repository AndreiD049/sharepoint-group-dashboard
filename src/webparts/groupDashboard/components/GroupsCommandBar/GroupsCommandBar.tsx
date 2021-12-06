import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { CommandBar } from "office-ui-fabric-react";
import * as React from "react";
import { FC } from "react";

export interface IGroupsCommandBarProps {
	selected: ISiteGroupInfo;
	setCreating: (boolean) => void;
	setEdited: (item: ISiteGroupInfo) => void;
	setDeleted: () => void;
}

const GroupsCommandBar: FC<IGroupsCommandBarProps> = (props) => {
	return (
		<CommandBar
			items={[
				{
					key: "createnew",
					text: "Create new",
					iconProps: {
						iconName: "Add",
					},
					onClick: () => props.setCreating(true),
				},
				{
					key: "edit",
					text: "Edit",
					iconProps: {
						iconName: "Edit",
					},
					disabled: !props.selected,
					onClick: () => props.setEdited(props.selected),
				},
				{
					key: "delete",
					text: "Delete",
					iconProps: {
						iconName: "Delete",
					},
					disabled: !props.selected,
					onClick: props.setDeleted,
				}
			]}
		/>
	);
};

export default GroupsCommandBar;
