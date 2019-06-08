import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'InjectJsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'InjectJsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IInjectJsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  // testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class InjectJsApplicationCustomizer
  extends BaseApplicationCustomizer<IInjectJsApplicationCustomizerProperties> {

  private _externalJs: string = "https://4698.sharepoint.com/sites/TeamSiteModern/SiteAssets/test.js";

  @override
  public onInit(): Promise<void> {
    console.log('InjectJsApplicationCustomizer.onInit(): Entered.');

    let scriptTag: HTMLScriptElement = document.createElement("script");
    scriptTag.src = this._externalJs;
    scriptTag.type = "text/javascript";

    document.getElementsByTagName("head")[0].appendChild(scriptTag);

    console.log('InjectJsApplicationCustomizer.onInit(): Added script link.');
    console.log('InjectJsApplicationCustomizer.onInit(): Leaving.');

    return Promise.resolve();
  }
}
