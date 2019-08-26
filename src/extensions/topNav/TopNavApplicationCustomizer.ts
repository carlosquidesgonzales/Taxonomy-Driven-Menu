import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import {
  ITerm,
  ITermData,
  Session,
  ITermStore,
  taxonomy,
  ITermGroup,
  ITermGroupData,
  ITermSet
} from "@pnp/sp-taxonomy";

import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'TopNavApplicationCustomizerStrings';
import TopNavBar, { ITopNavBarProps } from './components/TopNavBar';
import FooterNavBar, { IFooterNavBarProps } from './components/FooterNavBar';
import pnp from "sp-pnp-js";
import * as SPTermStore from './components/services/SPTermStoreService';

const LOG_SOURCE: string = 'TopNavApplicationCustomizer';
const NAV_TERMS_KEY: string = 'global-navigation-terms';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITopNavApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
  TopMenuTermSet?: string;
  BottomMenuTermSet?: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TopNavApplicationCustomizer
  extends BaseApplicationCustomizer<ITopNavApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _topMenuItems: SPTermStore.ISPTermObject[];
  private _bottomMenuItems: SPTermStore.ISPTermObject[];

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Configure caching
    pnp.setup({
      defaultCachingStore: "session",
      defaultCachingTimeoutSeconds: 900, //15min
      globalCacheDisable: false // true to disable caching in case of debugging/testing
    });

    const taxonomy = new Session(this.context.pageContext.web.absoluteUrl);
    const getTaxonomy = await taxonomy.termStores.get();
    const taxonomyName = getTaxonomy[0].Name;
    const stores: ITermStore = taxonomy.termStores.getByName(taxonomyName);

    // Retrieve the menu items from taxonomy
    let termStoreService: SPTermStore.SPTermStoreService = new SPTermStore.SPTermStoreService({
      siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      myTaxonomy: stores
    });

    console.log(this.context.pageContext.web.language);
    console.log(this.properties.TopMenuTermSet);
    // if (this.properties.TopMenuTermSet != undefined) {
    this._topMenuItems = await termStoreService.getTermsFromTermSetAsync(this.properties.TopMenuTermSet, this.context.pageContext.web.absoluteUrl);
    // }


    // if (this.properties.BottomMenuTermSet != undefined) {
    //   this._bottomMenuItems = await termStoreService.getTermsFromTermSetAsync(this.properties.BottomMenuTermSet, this.context.pageContext.web.language);
    // }
    // Call render method for generating the needed html elements
    // this._renderPlaceHolders();

    // return Promise.resolve<void>();

    // // Wait for the placeholders to be created (or handle them being changed) and then
    // // render.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {

    console.log('Available placeholders: ',
      this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose }
        );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this._topMenuItems != null && this._topMenuItems.length > 0) {
        const element: React.ReactElement<ITopNavBarProps> = React.createElement(
          TopNavBar,
          {
            menuItems: this._topMenuItems,
          }
        );

        ReactDom.render(element, this._topPlaceholder.domElement);
      }
    }

    // // Handling the bottom placeholder
    // if (!this._bottomPlaceholder) {
    //   this._bottomPlaceholder =
    //     this.context.placeholderProvider.tryCreateContent(
    //       PlaceholderName.Bottom,
    //       { onDispose: this._onDispose });

    //   // The extension should not assume that the expected placeholder is available.
    //   if (!this._bottomPlaceholder) {
    //     console.error('The expected placeholder (Bottom) was not found.');
    //     return;
    //   }

    //   if (this._bottomMenuItems != null && this._bottomMenuItems.length > 0) {
    //     const element: React.ReactElement<IFooterNavBarProps> = React.createElement(
    //       FooterNavBar,
    //       {
    //         menuItems: this._bottomMenuItems,
    //       }
    //     );

    //     ReactDom.render(element, this._bottomPlaceholder.domElement);
    //   }
    // }
  }

  private _onDispose(): void {
    console.log('[TenantGlobalNavBarApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
