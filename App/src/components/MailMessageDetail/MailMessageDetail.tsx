import * as React from 'react';
import styles from './MailMessageDetail.module.scss';
import { IMailMessageDetailProps } from './IMailMessageDetailProps';
import { isUndefined } from 'lodash';
import { IBaseComponentState } from '../../models';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Stack, TextField } from 'office-ui-fabric-react';

export const MailMessageDetail: React.FunctionComponent<IMailMessageDetailProps> = (props: IMailMessageDetailProps) => {
  const [state, setState] = React.useState<IBaseComponentState>({
    isLoading: false,
    hasError: false,
    errorMessage: ""
  });

  React.useEffect(() => {
    (async () => {
      if (isUndefined(props.message)) {
        return;
      }
      try {
        stateRef.current = {
          ...stateRef.current,
          isLoading: true,
        };
        setState(stateRef.current);
        // ...
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
  }, [props.message]);

  const stateRef = React.useRef(state); // Use to access state on eventListenners

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
      props.message ? (
        <>
          <div className={styles.mailmessagedetail}>
            <h3 className={styles.title}>Message details: {props.message.subject}</h3>
            <div className={styles.container}>
              <div className={styles.row}>
                <Stack>
                  <TextField label="Subject:" readOnly value={props.message.subject}/>
                  <TextField label="From:" readOnly value={props.message.from} />
                  <TextField label="Cc:" readOnly value={props.message.cc} />
                  <TextField label="To:"  readOnly value={props.message.to} />
                  <TextField label="Date:" readOnly value={props.message.date}/>
                  <TextField label="Body (text):" multiline rows={10} readOnly value={props.message.textBody} />
                </Stack>
              </div>
            </div>
          </div>
        </>
      ) : (
        <>
          <div className={styles.mailmessagedetail}>
            <div className={styles.container}>
              <div className={styles.row}>
                <div>
                  {'Select message to see more details ...'}
                </div>
              </div>
            </div>
          </div>
        </>
      )
    );
  }
}