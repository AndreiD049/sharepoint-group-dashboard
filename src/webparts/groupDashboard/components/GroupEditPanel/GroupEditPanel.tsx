import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import {
  Checkbox,
  DefaultButton,
  List,
  Panel,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
  Text,
  getTheme,
} from "office-ui-fabric-react";
import * as React from "react";
import styles from "./GroupEditedPanel.module.scss";
import { FC } from "react";
import {
  addUserToGroup,
  deleteUserFromGroup,
  getGroupUsers,
  getUsers,
  updateGroup,
} from "../../dal/groups";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";

export interface IGroupEditPanelProps {
  group: ISiteGroupInfo;
  handleEdit: (id: number, group: ISiteGroupInfo) => void;
  setEdited: (item: ISiteGroupInfo) => void;
  isOpen: boolean;
}

export interface IGroupEditFormValues {
  Title: string;
  Description: string;
}

const GroupEditPanel: FC<IGroupEditPanelProps> = (props) => {
  const theme = getTheme();
  const [groupUsers, setGroupUsers] = React.useState<{
    [key: string]: ISiteUserInfo;
  }>({});
  const [users, setUsers] = React.useState<ISiteUserInfo[]>([]);
  const [values, setValues] = React.useState<IGroupEditFormValues>({
    Title: props.group?.Title ?? "",
    Description: props.group?.Description ?? "",
  });
  const [search, setSearch] = React.useState<string>("");
  const [membersOnly, setMembersOnly] = React.useState<boolean>(false);

  // Update text field
  React.useEffect(() => {
    setValues((prev) => ({
      ...prev,
      Title: props.group?.Title,
      Description: props.group?.Description,
    }));
  }, [props.group]);

  // Load users
  React.useEffect(() => {
    async function run() {
      if (props.group) {
        const [groupUserList, siteUserList] = await Promise.all([
          getGroupUsers(props.group.Id),
          getUsers(),
        ]);
        const result = {};
        groupUserList.forEach((user) => (result[user.Id.toString()] = user));
        setGroupUsers(result);
        // Get only users that have an email
        setUsers(siteUserList.filter((user) => user.Email !== ""));
      } else {
        setGroupUsers({});
        setUsers([]);
      }
    }
    run();
  }, [props.group]);

  const isUserInGroup = (user: ISiteUserInfo) => {
    return groupUsers.hasOwnProperty(user.Id.toString());
  };

  const filterUsers = React.useCallback(() => {
    return users
      .filter((user) => (membersOnly ? isUserInGroup(user) : user))
      .filter(
        (user) => user.Title.toLowerCase().indexOf(search.toLowerCase()) !== -1
      );
  }, [search, users, membersOnly]);

  const handleCheck = async (user: ISiteUserInfo) => {
    if (isUserInGroup(user)) {
      // Remove user from the group
      await deleteUserFromGroup(user.Email, props.group.Id);
      // Remove the user from client list
      setGroupUsers((prev) => {
        const copy = { ...prev };
        delete copy[user.Id.toString()];
        return copy;
      });
      // Force update users
      setUsers((prev) => prev.slice());
    } else {
      // Add user to the group and to the list
      const addedUser = await addUserToGroup(user.LoginName, props.group.Id);
      // Add user to the list
      setGroupUsers((prev) => ({
        ...prev,
        [addedUser.Id.toString()]: addedUser,
      }));
      // Force update users
      setUsers((prev) => prev.slice());
    }
  };

  const handleChange =
    (key: keyof IGroupEditFormValues) =>
    (e: React.FormEvent, newValue: string) =>
      setValues((prev) => ({
        ...prev,
        [key]: newValue,
      }));

  const onDismiss = () => {
    props.setEdited(null);
    setSearch("");
    setMembersOnly(false);
  };

  const handleEdit = async () => {
    const updated = await updateGroup(props.group.Id, values);
    props.handleEdit(props.group.Id, updated);
    props.setEdited(null);
  };

  const footer = (
    <div className={styles.footer}>
      <PrimaryButton text="Update" onClick={handleEdit} />
      <DefaultButton text="Cancel" onClick={onDismiss} />
    </div>
  );

  return (
    <Panel
      headerText={`Edit '${props.group?.Title}'`}
      isOpen={props.isOpen}
      onDismiss={onDismiss}
      isLightDismiss
      isFooterAtBottom
      onRenderFooter={() => footer}
    >
      <TextField
        label="Name"
        value={values.Title}
        onChange={handleChange("Title")}
      />
      <TextField
        label="Description"
        value={values.Description}
        onChange={handleChange("Description")}
      />
      <Separator
        styles={{
          root: {
            marginTop: theme.spacing.m,
          },
        }}
      >
        Users
      </Separator>
      <Stack
        verticalAlign="space-evenly"
        horizontal
        styles={{
          root: {
            margin: `${theme.spacing.l1} 0`,
          },
        }}
      >
        <TextField
          placeholder="Search"
          value={search}
          styles={{
            root: {
              paddingRight: theme.spacing.m,
            },
          }}
          onChange={(e: any, newvalue: string) => setSearch(newvalue)}
        />
        <Checkbox
          label="Members only"
          checked={membersOnly}
          onChange={(e: any, checked) => setMembersOnly(checked)}
        />
      </Stack>
      <List
        items={filterUsers()}
        getKey={(item) => item.Id.toString()}
        onRenderCell={(item) => (
          <Stack horizontal horizontalAlign="start" verticalAlign="center">
            <Checkbox
              checked={isUserInGroup(item)}
              onChange={() => handleCheck(item)}
              onRenderLabel={() => (
                <div
                  style={{
                    marginBottom: theme.spacing.s1,
                    marginLeft: theme.spacing.s1,
                  }}
                >
                  <Text variant="smallPlus" block>
                    {item.Title}
                  </Text>
                  <Text variant="xSmall">{item.Email}</Text>
                </div>
              )}
            />
          </Stack>
        )}
      />
    </Panel>
  );
};

export default GroupEditPanel;
