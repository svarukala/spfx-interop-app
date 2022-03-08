import * as React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { useState, useEffect } from 'react';
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import { Configuration, LogLevel, PublicClientApplication, AccountInfo, SilentRequest, 
    InteractionRequiredAuthError, AuthorizationUrlRequest } from "@azure/msal-browser";
import { Providers, SharePointProvider, SimpleProvider, ProviderState } from '@microsoft/mgt-spfx';    
import SPOReusable from './SPOReusable';
import MSGReusable from './MSGReusable';
import MGTReusable from './MGTReusable';
import ShowAdaptiveCard from './ShowAdaptiveCard';
import {  FileList, PeoplePicker, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

const msalConfig: Configuration = {
  auth: {
    clientId: "c613e0d1-161d-4ea0-9db4-0f11eeabc2fd",
    authority: "https://login.microsoftonline.com/044f7a81-1422-4b3d-8f68-3001456e6406",
    redirectUri:"https://m365x229910.sharepoint.com/sites/DevDemo/_layouts/15/workbench.aspx",
  },
  cache: {
    cacheLocation: "localStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
},
  system: {
    iframeHashTimeout: 10000,
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};

const msalInstance: PublicClientApplication = new PublicClientApplication(
    msalConfig
  );

let currentAccount: AccountInfo = null;

const tokenrequest: SilentRequest = {
    scopes: ['Mail.Read','calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
    account: currentAccount,
}; 


function AppInit(props) {
  // Declare a new state variable, which we'll call "count"
  const [ssoToken, setSsoToken] = useState<string>();
  const [loginName, setLoginName] = useState<string>();
  const [error, setError] = useState<string>();
  const [accessToken, setAccessToken] = useState<string>();

  useEffect(() => {    
    if (props.loginName) {
        setLoginName(props.loginName);
    }

    if(ssoToken) { 
        //no action required
    }
    else { 
        getAccessTokenNonAsync();
    }

    if (!Providers.globalProvider) {
        console.log('Initializing global provider');
        Providers.globalProvider = new SimpleProvider(async ()=>{return getAccessToken()});  //new SharePointProvider(props.spoContext);
        Providers.globalProvider.setState(ProviderState.SignedIn);
    } 
  }, []);    

  const setCurrentAccount = (): void => {
    const currentAccounts: AccountInfo[] = msalInstance.getAllAccounts();
    if (currentAccounts === null || currentAccounts.length == 0) {
      currentAccount = msalInstance.getAccountByUsername(
        //this.context.pageContext.user.loginName
        loginName
      );
    } else if (currentAccounts.length > 1) {
      console.warn("Multiple accounts detected.");
      currentAccount = msalInstance.getAccountByUsername(
        //this.context.pageContext.user.loginName
        loginName
      );
    } else if (currentAccounts.length === 1) {
      currentAccount = currentAccounts[0];
    }
    tokenrequest.account = currentAccount;
  }; 

  const getAccessToken = async (): Promise<string> => {
    console.log("Getting access token async");
    let accessToken: string = null;
    setCurrentAccount();
    console.log(currentAccount);
    return msalInstance
      .acquireTokenSilent(tokenrequest)
      .then((tokenResponse) => {
        console.log("Inside Silent");
        console.log("Access token: "+ tokenResponse.accessToken);
        console.log("ID token: "+ tokenResponse.idToken);
        return tokenResponse.accessToken;
      })
      .catch((err) => {
        console.log(err);
        console.log("Silent Failed");
        if (err instanceof InteractionRequiredAuthError) {
          return interactionRequired();
        } else {
          console.log("Some other error. Inside SSO.");
          //const loginPopupRequest: AuthorizationUrlRequest = tokenrequest;
          const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
          loginPopupRequest.loginHint = loginName;
          return msalInstance
            .ssoSilent(loginPopupRequest)
            .then((tokenResponse) => {
              return tokenResponse.accessToken;
            })
            .catch((ssoerror) => {
              console.error(ssoerror);
              console.error("SSO Failed");
              if (ssoerror) {
                return interactionRequired();
              }
              return null;
            });
        }
      });
  };

  const getAccessTokenNonAsync = (): void => {
    console.log("Getting access token");
    let accessToken: string = null;
    setCurrentAccount();
    console.log(currentAccount);
    msalInstance
      .acquireTokenSilent(tokenrequest)
      .then((tokenResponse) => {
        console.log("Inside Silent");
        console.log("Access token: "+ tokenResponse.accessToken);
        console.log("ID token: "+ tokenResponse.idToken);
        setSsoToken(tokenResponse.idToken);
        setAccessToken(tokenResponse.accessToken);
      })
      .catch((err) => {
        console.log(err);
        console.log("Silent Failed");
        if (err instanceof InteractionRequiredAuthError) {
          interactionRequired();
        } else {
          console.log("Some other error. Inside SSO.");
          //const loginPopupRequest: AuthorizationUrlRequest = tokenrequest;
          const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
          loginPopupRequest.loginHint = loginName;
          return msalInstance
            .ssoSilent(loginPopupRequest)
            .then((tokenResponse) => {
                setSsoToken(tokenResponse.idToken);
                setAccessToken(tokenResponse.accessToken);
            })
            .catch((ssoerror) => {
              console.error(ssoerror);
              console.error("SSO Failed");
              if (ssoerror) {
                interactionRequired();
              }
              //return null;
              setError("SSO Failed");
            });
        }
      });
  };

  const reuseAccessToken = async (): Promise<string> => {
    return accessToken;
  };

  const interactionRequired = (): Promise<string> => {
    console.log("Inside Interaction");
    const loginPopupRequest: AuthorizationUrlRequest = tokenrequest as AuthorizationUrlRequest;
    loginPopupRequest.loginHint = loginName; //this.context.pageContext.user.loginName;
    return msalInstance
      .acquireTokenPopup(loginPopupRequest)
      .then((tokenResponse) => {
        //return tokenResponse.accessToken;
        setSsoToken(tokenResponse.idToken);
        setAccessToken(tokenResponse.accessToken);
      })
      .catch((error) => {
        console.error(error);
        // I haven't implemented redirect but it is fairly easy
        console.error("Maybe it is a popup blocked error. Implement Redirect");
        return null;
      });
  }; 

  const SiteResult = (props: MgtTemplateProps) => {
    const site = props.dataContext as MicrosoftGraph.Site;

    return (
        <div>
            <h1>{site.name}</h1>
            {site.webUrl}
      </div>
      );
    };

  return (
    <div>
        {error && "Error: " + error}
        {
            ssoToken &&
          
            <Pivot aria-label="Basic Pivot Example">
                <PivotItem headerText="SPO REST API">
                    <SPOReusable idToken={ssoToken} />
                </PivotItem>
                <PivotItem headerText="MS Graph REST API">
                    <MSGReusable idToken={ssoToken} />
                </PivotItem>
                <PivotItem headerText="MS Graph Toolkit">
                    <Pivot>
                        <PivotItem headerText="Files">
                            <FileList></FileList> 
                        </PivotItem>
                        <PivotItem headerText="People">
                            <br/>
                            <PeoplePicker></PeoplePicker>
                        </PivotItem>
                        <PivotItem headerText="File Upload">
                            <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
                        itemPath="/" enableFileUpload></FileList>
                        </PivotItem>
                        <PivotItem headerText="Sites Search Using MSGraph">
                            <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                                    <SiteResult template="value" />
                            </Get>
                        </PivotItem>
                    </Pivot>
                </PivotItem>
                <PivotItem headerText="Adaptive Card">
                    <ShowAdaptiveCard />
                </PivotItem>                
            </Pivot>
        }
    </div>
  );
}

export default AppInit;

/*
                    <Pivot>
                        <PivotItem headerText="Files">
                            <FileList></FileList> 
                        </PivotItem>
                        <PivotItem headerText="People">
                            <br/>
                            <PeoplePicker></PeoplePicker>
                        </PivotItem>
                        <PivotItem headerText="File Upload">
                            <FileList driveId="b!mKw3q1anF0C5DyDiqHKMr8iJr_oIRjlGl4854HhHtho07AdbOeaLT5rMH83yt89B" 
                        itemPath="/" enableFileUpload></FileList>
                        </PivotItem>
                        <PivotItem headerText="Sites Search Using MSGraph">
                            <Get resource="/sites?search=contoso" scopes={['Sites.Read.All']} maxPages={2}>
                                    <SiteResult template="value" />
                            </Get>
                        </PivotItem>
                    </Pivot>
*/