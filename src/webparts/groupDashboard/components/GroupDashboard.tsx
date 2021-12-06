import * as React from "react";
import styles from "./GroupDashboard.module.scss";
import './styles/global.css';
import { IGroupDashboardProps } from "./IGroupDashboardProps";
import { FC } from "react";
import {
  DetailsList,
  SelectionMode,
  Text,
  Selection,
  DetailsListLayoutMode,
} from "office-ui-fabric-react";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { deleteGroup, getGroups } from "../dal/groups";
import GroupsCommandBar from "./GroupsCommandBar";
import GroupCreatePanel from "./GroupCreatePanel/GroupCreatePanel";
import GroupEditPanel from "./GroupEditPanel";

const GroupDashboard: FC<IGroupDashboardProps> = (props) => {
  const [groups, setGroups] = React.useState<ISiteGroupInfo[]>([]);
  const [selected, setSelected] = React.useState<ISiteGroupInfo>(null);
  const [editedGroup, setEditedGroup] = React.useState<ISiteGroupInfo>(null);
  const [isCreating, setIsCreating] = React.useState<boolean>(false);
  const selection = new Selection({
    getKey: (item: ISiteGroupInfo) => item.Id,
    selectionMode: SelectionMode.single,
    onSelectionChanged: () =>
      setSelected(
        selection.getSelectedCount() > 0 ? selection.getSelection()[0] : null
      ),
  });

  React.useEffect(() => {
    async function run() {
      setGroups(await getGroups());
    }
    run();
  }, []);

  const handleGroupCreated = (group: ISiteGroupInfo) => {
    setGroups((prev) => [...prev, group]);
  };

  const handleGroupEdit = (id: number, updatedGroup: ISiteGroupInfo) => {
    setGroups(prev => prev.map(group => group.Id === id ? updatedGroup : group));
  };

  const handleGroupDeleted = async () => {
    const id = selected.Id;
    await deleteGroup(selected.Id);
    setGroups(prev => prev.filter(group => group.Id !== id));
  };

  return (
    <div>
      <Text variant="xLarge">{props.description ?? 'Group Dashboard'}</Text>
      <GroupsCommandBar
        selected={selected}
        setCreating={setIsCreating}
        setEdited={setEditedGroup}
        setDeleted={handleGroupDeleted}
      />
      <div>
        <DetailsList
          selection={selection}
          selectionMode={SelectionMode.single}
          layoutMode={DetailsListLayoutMode.justified}
          columns={[
            {
              key: "title",
              name: "Title",
              fieldName: "Title",
              minWidth: 150,
              maxWidth: 200,
              isResizable: true,
            },
            {
              key: "description",
              name: "Description",
              fieldName: "Description",
              minWidth: 200,
              isResizable: true,
            }
          ]}
          items={groups}
          onItemInvoked={(item) => setEditedGroup(item)}
        />
      </div>
      <GroupEditPanel
        group={editedGroup}
        handleEdit={handleGroupEdit}
        isOpen={editedGroup !== null}
        setEdited={setEditedGroup}
      />
      <GroupCreatePanel
        isOpen={isCreating}
        setIsOpen={setIsCreating}
        handleGroupCreated={handleGroupCreated}
      />
    </div>
  );
};

export default GroupDashboard;
