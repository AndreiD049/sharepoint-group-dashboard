import { PrimaryButton } from "@microsoft/office-ui-fabric-react-bundle";
import { ISiteGroupInfo } from "@pnp/sp/site-groups";
import { DefaultButton, Panel, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { FC } from "react";
import { createGroup } from "../../dal/groups";
import styles from "./GroupCreatePanel.module.scss";

export interface IGroupCreatePanelProps {
  isOpen: boolean;
  setIsOpen: (boolean) => void;
	handleGroupCreated: (ISiteGroupInfo) => void;
}

export interface IGroupCreateFormValues {
  Title: string;
	Description: string;
}

const GroupCreatePanel: FC<IGroupCreatePanelProps> = (props) => {
  const [values, setValues] = React.useState<IGroupCreateFormValues>({
    Title: "",
		Description: "",
  });

	const panelDismiss = () => {
		props.setIsOpen(false);
		setValues({
			Title: "",
			Description: "",
		});
	};
	
	const handleCreate = async () => {
		try {
			const group = await createGroup(values);
			props.handleGroupCreated(group);
			panelDismiss();
		} catch (e) {
			console.error(e);
		}
	};


	const footer = (
		<div className={styles.footer}>
			<PrimaryButton text="Create" onClick={handleCreate} />
			<DefaultButton text="Cancel" onClick={() => props.setIsOpen(false)} />
		</div>
	);

  const handleInput =
    (key: keyof IGroupCreateFormValues) =>
    (e: React.FormEvent, newValue: string) =>
      setValues((prev) => ({
        ...prev,
        [key]: newValue,
      }));

  return (
    <Panel
      isOpen={props.isOpen}
      headerText="Create group"
      onDismiss={panelDismiss}
      isFooterAtBottom
			onRenderFooter={() => footer}
    >
      <TextField
        label="Name"
        value={values.Title}
        onChange={handleInput("Title")}
      />
			<TextField
				label="Description"
				value={values.Description}
				onChange={handleInput("Description")}
			/>
    </Panel>
  );
};

export default GroupCreatePanel;
