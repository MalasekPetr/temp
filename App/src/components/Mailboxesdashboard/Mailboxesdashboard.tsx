import * as React from 'react';
import styles from './Mailboxesdashboard.module.scss';
import { IMailboxesdashboardProps } from './IMailboxesdashboardProps';
import { IMailboxesdashboardState } from './IMailboxesdashboardState';
import { isUndefined } from 'lodash';
import {
  Tile
} from '../../components';
import { IBaseComponentState, IMailBoxApp} from '../../models';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Stack } from 'office-ui-fabric-react';
import { useMailHooks } from '../../hooks';

export const Mailboxesdashboard: React.FunctionComponent<IMailboxesdashboardProps> = (props: IMailboxesdashboardProps) => {
  const { getMailBoxApps } = useMailHooks();
  
  const [state, setState] = React.useState<IBaseComponentState>({
    isLoading: true,
    hasError: false,
    errorMessage: ""
  });

  const [apps, setApps] = React.useState([]);

  React.useEffect(() => {
    (async () => {
      if (isUndefined(props.webpartprops.backendapi)) {
        return;
      }
      try {
        stateRef.current = {
          ...stateRef.current,
          isLoading: true,
        };
        setState(stateRef.current);
        setApps(await getMailBoxApps(props.webpartprops.backendapi));
        stateRef.current = {
          ...stateRef.current,
          isLoading: false,
        };
        setState(stateRef.current);
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
  }, []);

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
      <>
        <div className={styles.mailboxesdashboard}>
          <h1 className={styles.title}>{props.webpartprops.title}</h1>
          <span className={styles.description}>{props.webpartprops.description}</span>
          <div className={styles.container}>
            <div className={styles.row}>
              <div>
                {apps.map((m: IMailBoxApp, i: number) => {
                  m.backendapi = props.webpartprops.backendapi;
                  return (<Tile key={i} app={m} interval={props.webpartprops.refreshinterval} />);
                })}
              </div>
            </div>
          </div>
        </div>
      </>
    );
  }
}