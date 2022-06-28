import * as React from 'react';
require('../styles/App.css');
import { ITileProps, ITileState } from '.';
import { isUndefined } from 'lodash';
import { MessageBar, MessageBarType, Spinner, SpinnerSize, Stack } from 'office-ui-fabric-react';
import { useMailHooks } from '../../hooks';
import { IBaseComponentState } from '../../models';

export const Tile: React.FunctionComponent<ITileProps> = (props: ITileProps) => {
  const { countNewItems } = useMailHooks();

  const [nrofitems, setNrofitems] = React.useState(0);
 
  const [state, setState] = React.useState<IBaseComponentState>({
    isLoading: true,
    hasError: false,
    errorMessage: ""
  });

  async function load() {
    stateRef.current = {
      ...stateRef.current,
      isLoading: true,
    };
    setState(stateRef.current);
    setNrofitems(await countNewItems(props.app));
    stateRef.current = {
      ...stateRef.current,
      isLoading: false,
    };
    setState(stateRef.current);
  }

  //const addresswithshy = props.app.address.replace('@','&#173;@')
  
  React.useEffect(() => {
    (async () => {
      if (isUndefined(props.app)) {
        return;
      }
      try {
        await load();
          const timeout = setInterval(async (): Promise<void> => {
            await load();
          }, props.interval);
          return () => clearInterval(timeout);
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
        <div className='mboxapp-tile'>
          <h3>{props.app.name}</h3>
          <span>
            {<a href={`${props.app.spWebBaseUrl}${props.app.appAddress}`}>{props.app.address}</a>}
          </span><br />
          <span>New Messages: {nrofitems}</span>
        </div>
      </>
    );
  }
}