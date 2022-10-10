import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { HomeComponent } from './home/home.component';

import { MsalGuardConfiguration, MsalModule, MsalGuard, MsalInterceptorConfiguration } from '@azure/msal-angular';
import { Configuration, PublicClientApplication, InteractionType  } from '@azure/msal-browser';
import { LoginComponent } from './login/login.component';

const isIE = window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;

const AZURE_CONSTANTS = {
  applicationID: 'f0419ec1-e7ee-4db7-9d4f-c85559f87692',
  cloudInstanceID: 'https://login.microsoftonline.com',
  tenantInfo: '5a4863ed-40c8-4fd5-8298-fbfdb7f13095',
  redirectURI: 'http://localhost:4200/home',
};

const msalConfig: Configuration = {
  auth: {
    clientId: AZURE_CONSTANTS.applicationID, // This is your client ID
    authority: `${AZURE_CONSTANTS.cloudInstanceID}/${AZURE_CONSTANTS.tenantInfo}`, // This is your tenant ID
    redirectUri: AZURE_CONSTANTS.redirectURI// This is your redirect URI
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: isIE, // Set to true for Internet Explorer 11
  }
};

const msalGuardConfig: MsalGuardConfiguration = {
  interactionType: InteractionType.Popup, // MSAL Guard Configuration
  authRequest: {
    scopes: ['user.read']
  }
};

const msalInterceptor: MsalInterceptorConfiguration = {
  interactionType: InteractionType.Redirect, // MSAL Interceptor Configuration
  protectedResourceMap: new Map([ 
    ['Enter_the_Graph_Endpoint_Here/v1.0/me', ['user.read']]
  ])
};

@NgModule({
  declarations: [
    AppComponent,
    HomeComponent,
    LoginComponent,
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    MsalModule.forRoot(new PublicClientApplication(msalConfig), msalGuardConfig, msalInterceptor)
  ],
  providers: [MsalGuard],
  bootstrap: [AppComponent]
})
export class AppModule { }