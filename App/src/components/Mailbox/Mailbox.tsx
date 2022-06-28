import * as React from 'react';
import styles from './Mailbox.module.scss';
import { IMailboxProps } from './IMailboxProps';
import { IMailboxState, MailThread } from '../../components';
import { IMessage} from '../../models';
import { useMailHooks } from '../../hooks';
import { isUndefined } from 'lodash';
import { 
  MessageBar, 
  MessageBarType, 
  Spinner, 
  SpinnerSize, 
  Stack, 
  CommandBar, 
  ICommandBarItemProps, 
  MarqueeSelection,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
  Separator
} from 'office-ui-fabric-react';
import { MailMessageDetail } from '../MailMessageDetail/MailMessageDetail';
import { IItems } from '@pnp/sp/items';

export const Mailbox: React.FunctionComponent<IMailboxProps> = (props: IMailboxProps) => {
  const { getAllItems, convertIItems, sendAutoResponse } = useMailHooks();

  const [state, setState] = React.useState<IMailboxState>({
    isLoading: true,
    hasError: false,
    errorMessage: "",
    selectionDetails: getSelectionDetails()
  });

  const [messages, setMessages] = React.useState<IMessage[]>([] as IMessage[]);
  const [selectedId, setSelectedId] = React.useState<number>(0);
  const [selection, setSelection] = React.useState<Selection | undefined>();

  const _selection: Selection = new Selection({
    onSelectionChanged: () => {
      stateRef.current = {
        ...stateRef.current,
        selectionDetails: getSelectionDetails()
      };
      setState(stateRef.current);
    }
  });

  React.useEffect(() => {
    (async () => {
      if (isUndefined(props.address)) {
        return;
      }
      try {
        stateRef.current = {
          ...stateRef.current,
          isLoading: true,
        };
        setState(stateRef.current);
        setSelection(_selection);
        setMessages(convertIItems(await getAllItems(props.spWebBaseUrl, props.spListId)));
        stateRef.current = {
          ...stateRef.current,
          isLoading: false,
        };
        setState(stateRef.current);
      } catch (error) {
        console.error(error);
        stateRef.current = {
          ...stateRef.current,
          hasError: true,
          errorMessage: error.message,
        };
        setState(stateRef.current);
      }
    })();
  }, [props.spListId]);

  React.useEffect(() => {
    (async () => {
      if (isUndefined(props.address)) {
        return;
      }
      try {
        setSelectedId(getSelectedId());
      } catch (error) {
        console.error(error);
        stateRef.current = {
          ...stateRef.current,
          hasError: true,
          errorMessage: error.message,
        };
        setState(stateRef.current);
      }
    })();
  }, [_selection]);

  const stateRef = React.useRef(state); // Use to access state on eventListenners

  function getSelectedId(): number {
    let retval: number
    const selectionCount = selection ? selection.getSelectedCount() : 0;
    if (selectionCount === 1) {
      retval = (selection.getSelection()[0] as IMessage).id as number;
    }
    return retval;
  }

  function getSelectionDetails(): string {
    const selectionCount = selection ? selection.getSelectedCount() : 0;
    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:  
        return (
          "1 item selected: " +
          (selection.getSelection()[0] as IItems)['MessageSubject']
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  const columns: IColumn[] = [
    {
      key: "column1",
      name: "Subject",
      fieldName: "subject",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: "column2",
      name: "From",
      fieldName: "from",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true
    },
    {
      key: "column3",
      name: "Date",
      fieldName: "date",
      minWidth: 180,
      maxWidth: 180,
      isResizable: true
    }
  ];

  const commands: ICommandBarItemProps[] = [
    {
      key: 'confirm',
      text: 'Confirm (Automatic Reply)',
      iconProps: { iconName: 'MailReply' },
      onClick: () => {
        sendAutoResponse(props.backendapi, props.spWebBaseUrl, props.spListId, messages.filter(i => i.id === selectedId)[0])
      }
    } 
  ];

  // Show Error if Exists
  if (state.hasError) {
    return (
      <>
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {state.errorMessage}
        </MessageBar>
      </>
    );
  }
  // Show Spinner while loading
  if (state.isLoading) {
    return (
      <>
        <Stack horizontal horizontalAlign="center">
          <Spinner size={SpinnerSize.medium}></Spinner>
        </Stack>
      </>
    );
  }
  // Show data
  if (!state.isLoading) {
    return (
    selection ? (
      <>
        <div className={styles.mailbox}>
          <h2 className={styles.title}>MailBox app - {props.name}</h2>
          {/*
          <CommandBar
            items={commands}
            ariaLabel={'Use left and right arrow keys to navigate between commands'}
          />
          */}
          <div className={styles.container}>
            <div className={styles.row}>
                <MarqueeSelection selection={selection}>
                  <DetailsList
                    items={messages}
                    compact={true}
                    columns={columns}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.single}
                    selection={selection}
                    selectionPreservedOnEmptyClick={true}
                    ariaLabelForSelectionColumn="Toggle selection"
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                    checkButtonAriaLabel="select row"
                  />
                </MarqueeSelection>
            </div>
            <Separator />
            <div className={styles.row}>
              {<MailThread 
                threadid={ selectedId > 0 ? messages.filter(i => i.id === selectedId)[0].threadId : 0} 
                spWebBaseUrl={props.spWebBaseUrl} 
                spListId={props.spListId}
                backendapi={props.backendapi} />}
            </div>
            {/*
            <Separator />
            <div className={styles.row}>
              {<MailMessageDetail message={messages.filter(i => i.id === selectedId)[0]} />}
            </div>
            */}
          </div >
        </div >
      </>
    ) : (
      <>
        <Stack horizontal horizontalAlign="center">
          <Spinner size={SpinnerSize.medium}></Spinner>
        </Stack>
      </>
    )
  )}
}