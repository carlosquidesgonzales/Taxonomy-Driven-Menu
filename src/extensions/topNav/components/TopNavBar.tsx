import * as React from 'react';
import * as SPTermStore from './services/SPTermStoreService';
import styles from './AppCustomizer.module.scss';
import './MyCss.css'
import { Nav, INavLinkGroup } from 'office-ui-fabric-react/lib/Nav';


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
import { Icon } from 'office-ui-fabric-react/lib/Icon';

export interface ITopNavBarProps {
  menuItems: SPTermStore.ISPTermObject[];
}

export interface ITopNavBarState {
  sideDrawerOpen: boolean;
  menuItems: IContextualMenuItem[];
}

export default class TopNavBar extends React.Component<ITopNavBarProps, ITopNavBarState> {
  constructor() {
    super();
    this.state = {
      sideDrawerOpen: false,
      menuItems: []
    };
  }


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
      onClick: () => { window.open(menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"], '_self') },
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

  private menuItem(menuItem) {
    return ({
      links: [
        {
          name: menuItem.name,
          url: menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"],
          links: menuItem.terms.length > 0 ?
            menuItem.terms.map((i) => { return (this.submenus(i)); }) : null
        }
      ]
    });
  }

  private submenus(menuItem) {
    return ({
      name: menuItem.name,
      url: menuItem.localCustomProperties["_Sys_Nav_SimpleLinkUrl"],
      links: menuItem.terms.length > 0 ?
        menuItem.terms.map((i) => { return (this.submenus(i)); }) : null
    })
  }



  public drawerToggleClickHandler = () => {
    this.setState(prevState => {
      return { sideDrawerOpen: !this.state.sideDrawerOpen }
    })
  }

  public backdropClickHandler = () => {
    this.setState({ sideDrawerOpen: false })
  }

  public render(): React.ReactElement<ITopNavBarProps> {
    const commandBarItems: IContextualMenuItem[] = this.props.menuItems.map((i) => {
      // console.log(i);
      return (this.projectMenuItem(i, ContextualMenuItemType.Header));
    });

    const navItems: INavLinkGroup[] = this.props.menuItems.map((i) => {
      return this.menuItem(i);
    })

    let attachedClasses = [styles.SideDrawer, styles.Close];
    if (this.state.sideDrawerOpen) {
      attachedClasses = [styles.SideDrawer, styles.Open];
    }

    return (
      // <div>

      //   <div className={styles.MainToolbar}>
      //     <Nav
      //       className={styles.Nav}
      //       expandButtonAriaLabel="Expand or collapse"
      //       groups={navItems} />
      //   </div>

      //   {/* {Side Drawer} */}
      //   {/* <div className = {styles.MiniToolbar}> */}
      //   {this.state.sideDrawerOpen ?
      //     <div className={styles.ToolBar}>
      //       <div className={styles.Backdrop} onClick={this.backdropClickHandler.bind(this)}></div>
      //       <div className={this.state.sideDrawerOpen ? styles.Open : styles.Close}>

      //        <div className = { styles.CloseIcon} onClick={this.drawerToggleClickHandler.bind(this)}>
      //           <Icon iconName="ChromeClose" />
      //        </div>
      //         <Nav
      //           className={styles.Nav}
      //           expandButtonAriaLabel="Expand or collapse"
      //           groups={navItems} />

      //       </div>
      //     </div> : null}
      //   {/* </div> */}

      //   {/* {Button Toggle} */}
      //   <div className={styles.Hamburger}>
      //     <button className={styles.ToggleButton} onClick={this.drawerToggleClickHandler.bind(this)}>
      //       <div className={styles.ButtonLine} />
      //       <div className={styles.ButtonLine} />
      //       <div className={styles.ButtonLine} />
      //     </button>
      //   </div>

      // </div>

      <div>
        <CommandBar
            isSearchBoxVisible={ false }
            elipisisAriaLabel='More options'
            items={ commandBarItems }
            />
      </div>

      //   <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.app}`}>
      //   <div className={`ms-bgColor-neutralLighter ms-fontColor-white ${styles.top}`}>
      //       <CommandBar
      //       className={styles.commandBar}
      //       isSearchBoxVisible={ false }
      //       elipisisAriaLabel='More options'
      //       items={ commandBarItems }
      //       />
      //   </div>
      // </div>
    );
  }
}
