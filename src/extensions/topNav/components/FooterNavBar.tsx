import * as React from 'react';
import * as SPTermStore from './services/SPTermStoreService';
import styles from './AppCustomizer.module.scss';
// import { sp, Web } from "@pnp/sp";
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

import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem, ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';

export interface IFooterNavBarProps {
  menuItems: SPTermStore.ISPTermObject[];
}

export interface IFooterNavBarState {

}

export default class FooterNavBar extends React.Component<IFooterNavBarProps, IFooterNavBarState> {
  constructor() {
    super();
    this.state = {
    };
  }

  // constructor(props: IFooterNavBarProps) {
  //   super(props);
  //   this.state = {
  //     menuItems: []
  //   };
  // }

  // public componentDidMount() {
  //   let id = this.props.termSetId;
  //   let url = this.props.siteUrl;

  //   this.getTerms(id, url).then(items => {
  //     let navItems = items.map(i => {
  //       return {
  //         href: i.LocalCustomProperties._Sys_Nav_SimpleLinkUrl,
  //         title: i.Name,
  //         name: i.Name
  //       } as ICommandBarItemProps;
  //     });
  //     this.setState({ terms: navItems });
  //     console.log(navItems);
  //   });
  // }

  // public async getTerms(groupId: string, siteUrl) {
  //   const taxonomy = new Session(siteUrl);
  //   const getTaxonomy = await taxonomy.termStores.get();
  //   const taxonomyName = getTaxonomy[0].Name;
  //   const store: ITermStore = taxonomy.termStores.getByName(taxonomyName);

  //   //Get group
  //   const group = await store.groups.get();
  //   let groupItemId;
  //   for (let index = 0; index < group.length; index++) {
  //     if (group[index].Name === "NavBar") {
  //       let itemId = group[index].Id;
  //       groupItemId = itemId.substring(itemId.indexOf('(') + 1, itemId.lastIndexOf(')'));
  //     }
  //   }

  //   // Get termsets in a group - termsets in group
  //   const grId: ITermGroup = await store.getTermGroupById(groupItemId); // GroupID
  //   let termsSets = await grId.termSets.get();
  //   let termSetId;
  //   for (let index = 0; index < termsSets.length; index++) {
  //     if (termsSets[index].Name === 'TopNavBar') {
  //       let itemId = termsSets[index].Id;
  //       termSetId = itemId.substring(itemId.indexOf('(') + 1, itemId.lastIndexOf(')'));
  //     }
  //   }

  //   //Get terms in a Termset  -terms in termsets
  //   const set: ITermSet = store.getTermSetById(termSetId); //TermsetsID
  //   const terms = await set.terms.get();

  //   return terms;
  // }

  private projectMenuItem(menuItem: SPTermStore.ISPTermObject, itemType: ContextualMenuItemType): IContextualMenuItem {
    return ({
      key: menuItem.identity,
      name: menuItem.name,
      itemType: itemType,
      href: menuItem.terms.length == 0 ?
        (menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"] != undefined ?
          menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]
          : null)
        : null,
      subMenuProps: menuItem.terms.length > 0 ?
        { items: menuItem.terms.map((i) => { return (this.projectMenuItem(i, ContextualMenuItemType.Normal)); }) }
        : null,
      isSubMenu: itemType != ContextualMenuItemType.Header,
    });
  }

  public render(): React.ReactElement<IFooterNavBarProps> {
    const commandBarItems: IContextualMenuItem[] = this.props.menuItems.map((i) => {
      console.log(i);
      return (this.projectMenuItem(i, ContextualMenuItemType.Header));
    });
    return (
      // <div>
      //   {/* <CommandBar
      //       isSearchBoxVisible={ false }
      //       elipisisAriaLabel='More options'
      //       items={ this.state.terms }
      //       /> */}
      // </div>

      <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.app}`}>
      <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.top}`}>
          <CommandBar
          className={styles.commandBar}
          isSearchBoxVisible={ false }
          elipisisAriaLabel='More options'
          items={ commandBarItems }
          />
      </div>
    </div>
    );
  }
}
