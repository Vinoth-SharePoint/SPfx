import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp/presets/all";
import styles from './Components/SPOSiteMessage.module.scss';
import * as strings from 'SPfxSiteMessageSolutionApplicationCustomizerStrings';
import SPOSiteMessageService,{ IBannerMessageItemProps } from './Components/SPOSiteMessageService';
const LOG_SOURCE: string = 'SPfxSiteMessageSolutionApplicationCustomizer';
export const QUALIFIED_NAME = 'Extension.ApplicationCustomizer.SPOSiteMessageService';

export interface ISPfxSiteMessageSolutionApplicationCustomizerProperties {
  listName: string;
}

export default class SPfxSiteMessageSolutionApplicationCustomizer
  extends BaseApplicationCustomizer<ISPfxSiteMessageSolutionApplicationCustomizerProperties> {

  @override
  protected async onInit(): Promise<void> {

    await super.onInit();
     sp.setup({
      spfxContext: this.context,
    });

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this.properties.listName) {
      const e: Error = new Error('Missing required configuration parameters');
      Log.error(QUALIFIED_NAME, e);
      return Promise.reject(e);
    }

    const header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

    if (!header) {
      `<div class ="${styles}"></div>`;
      const error = new Error('Could not find placeholder Top');
      Log.error(QUALIFIED_NAME, error);
      return Promise.reject(error);
     
    }

    const elem: React.ReactElement<IBannerMessageItemProps> = React.createElement(SPOSiteMessageService, {
      listName: this.properties.listName,

    });

    ReactDOM.render(elem, header.domElement);

    return Promise.resolve();

  }
}
