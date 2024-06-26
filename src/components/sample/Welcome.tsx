import { useContext, useEffect, useState } from "react";
import {
  Image,
  TabList,
  Tab,
  SelectTabEvent,
  SelectTabData,
  TabValue,
} from "@fluentui/react-components";
import "./Welcome.css";
import { useData } from "@microsoft/teamsfx-react";

import { TeamsFxContext } from "../Context";
import { app } from "@microsoft/teams-js";
import * as microsoftTeams from "@microsoft/teams-js";
import { Button } from "@fluentui/react-components";

interface Message {
  UUID: string | null;
}

interface MessageWithCallback {
  UUID: string | null;
  callbackId: string | null;
}

declare global {
  interface Window {
    authToken: (token: string) => void;
    webkit: {
      messageHandlers: {
        PageLoadCompleted: {
          postMessage: (message: Message) => void;
        };
        ExitWebView: {
          postMessage: (message: Message) => void;
        };
        AuthToken: {
          postMessage: (message: MessageWithCallback) => void;
        };
      };
    };
  }
}

export function Welcome(props: { showFunction?: boolean; environment?: string }) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";


  const [accessToken, setAccessToken] = useState("");
  const [timeTaken, setTimeTaken] = useState("");

  const { teamsUserCredential } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });
  
  const userName = loading || error ? "" : data!.displayName;
  const hubName = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    return context.app.host.name;
  })?.data;

  const getClientSideToken = (): Promise<string> => {

    return new Promise((resolve, reject) => {
      microsoftTeams.authentication.getAuthToken().then((result) => {
        setAccessToken(result);
         
        // Display the duration on the webpage
        // Assuming there's an element with the ID 'timeDisplay' to show the duration  
        resolve(result as string);
      }).catch((error) => {
        const errorMessage = "Error getting token: " + error;
        console.log(errorMessage);
        setAccessToken(errorMessage); // Note: This will also set an error message as the access token
        reject(errorMessage);
      });
    });
  };

  const buttonClicked = async() => {
    const startTime = performance.now();

    await getClientSideToken().then((result) => {
      const endTime = performance.now();
      const duration = endTime - startTime; 
      console.log(`setAccessToken took ${duration} milliseconds ${accessToken}`);
      setTimeTaken(`access token fetch took ${duration} milliseconds`);
    })
  }

  const [uuid, setUuid] = useState('');
  const [authToken, setAuthToken] = useState('Fetch Token Time');
  let callbackIndex: number = 1000;
  useEffect(() => {
    if (window.webkit && window.webkit.messageHandlers) {
      // Call your method here'
      const allParams = new URLSearchParams(window.location.search);
      const uuid = allParams.get('UUID') || '';
      setUuid(uuid);
      window.webkit.messageHandlers.PageLoadCompleted.postMessage({ UUID: uuid });
    }
  }, []);

  const handleClick = (): void => {
    window.webkit.messageHandlers.ExitWebView.postMessage({ UUID: uuid });
  };

  const handleAuthToken = (): void => {
    const startTime = Date.now();
    const data = new Promise<string>((resolve, reject) => {
      if (window.webkit && window.webkit.messageHandlers) {
        window.authToken = (token: string): void => {
          resolve(token);
        };
        window.webkit.messageHandlers.AuthToken.postMessage({
          UUID: uuid,
          callbackId: `${Date.now().toString()}-${callbackIndex++}`,
        });
      } else {
        reject('Interface not found');
      }
    });
    data.then(async (output) => {
      console.log(output);
      setAuthToken((Date.now() - startTime).toString());
    });
  };

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <h1 className="">Hi{userName ? ", " + userName : ""}</h1>
        {hubName && <p className="">This app is running in Hub: {hubName}</p>}
        <p className="">This app is running in: {friendlyEnvironmentName}</p>

        <Button appearance="primary" style={{ marginBottom: 10}} onClick={buttonClicked}>
          Authenticate
        </Button>

        <div>Auth Token</div>
        <div style={{ border: '1px solid #ccc', padding: '10px', maxWidth: '100%', margin: '0 auto', wordWrap: 'break-word' }}>
            {accessToken}
        </div>
        <div>
        {timeTaken}
        </div>
      </div>
      <div>
        <h1>Native Webview Method</h1>
        <button onClick={handleClick}>Click to Exit</button>
        <button onClick={handleAuthToken}>Get Auth Token</button>
      </div>
    </div>
  );
}
