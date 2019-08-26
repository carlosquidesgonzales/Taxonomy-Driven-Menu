import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
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

/**
 * @interface
 * Interface for SPTermStoreService configuration
 */
export interface ISPTermStoreServiceConfiguration {
  siteAbsoluteUrl: string;
  myTaxonomy: ITermStore;
}


/**
 * @interface
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
  identity: string;
  isAvailableForTagging: boolean;
  name: string;
  guid: string;
  customSortOrder: string;
  terms: ISPTermObject[];
  localCustomProperties: any;
}

/**
 * @class
 * Service implementation to manage term stores in SharePoint
 * Basic implementation taken from: https://oliviercc.github.io/sp-client-custom-fields/
 */
export class SPTermStoreService {
  private siteAbsoluteUrl: string;
  private myTaxonomy: ITermStore;

  /**
   * @function
   * Service constructor
   */
  constructor(config: ISPTermStoreServiceConfiguration) {
    this.siteAbsoluteUrl = config.siteAbsoluteUrl;
    this.myTaxonomy = config.myTaxonomy
  }

  /**
   * @function
   * Gets the collection of term stores in the current SharePoint env
   */
  public async getTermsFromTermSetAsync(termSetName: string, siteUrl): Promise<ISPTermObject[]> {

    if (Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint) {


      // const taxonomy = new Session(this.siteAbsoluteUrl);
      // const getTaxonomy = await taxonomy.termStores.get();
      // const taxonomyName = getTaxonomy[0].Name;
      // const store: ITermStore = taxonomy.termStores.getByName(taxonomyName);

      const store: ITermStore = this.myTaxonomy;

      //Get group
      const group = await store.groups.get();
      let groupItemId;
      for (let index = 0; index < group.length; index++) {
        if (group[index].Name === "NavBar") {
          let itemId = group[index].Id;
          groupItemId = itemId.substring(itemId.indexOf('(') + 1, itemId.lastIndexOf(')'));
        }
      }

      // Get termsets in a group - termsets in group
      const grId: ITermGroup = await store.getTermGroupById(groupItemId); // GroupID
      let termsSets = await grId.termSets.get();
      let termSetId;
      for (let index = 0; index < termsSets.length; index++) {
        if (termsSets[index].Name === termSetName) {
          let itemId = termsSets[index].Id;
          termSetId = itemId.substring(itemId.indexOf('(') + 1, itemId.lastIndexOf(')'));
        }
      }

      //Get terms in a Termset  -terms in termsets
      const set: ITermSet = store.getTermSetById(termSetId); //TermsetsID
      const terms = await set.terms.get();

      let item: (ITermData & ITerm)[] = await terms.filter(i => {
        // console.log(i.PathOfTerm.indexOf(';') > -1);
        // if(!(i.PathOfTerm.indexOf(';') > -1)){
        return !(i.PathOfTerm.indexOf(';') > -1);
        // }
      });


      return (await Promise.all<ISPTermObject>(item.map(async (t: any): Promise<ISPTermObject> => {
        return await this.projectTermAsync(t, this.siteAbsoluteUrl);
      })));
    }


    // // Default empty array in case of any missing data
    return (new Promise<Array<ISPTermObject>>((resolve, reject) => {
      resolve(new Array<ISPTermObject>());
    }));
  }


  /**
   * @function
   * Gets the child terms of another term of the Term Store in the current SharePoint env
   */
  private async getChildTermsAsync(term: any, siteUrl): Promise<ISPTermObject[]> {

    let termId = this.cleanGuid(term['Id']);

    // Check if there are child terms to search for
    if (Number(term['TermsCount']) > 0) {


      // const taxonomy = new Session(siteUrl);
      //   const getTaxonomy = await taxonomy.termStores.get();
      //   const taxonomyName = getTaxonomy[0].Name;
      //   const store: ITermStore = await taxonomy.termStores.getByName(taxonomyName);

      const store: ITermStore = this.myTaxonomy;
      // Get term in term
      const term: ITerm = await store.getTermById(termId); //TermsetsID
      // load the data into the terms instances
      const term2: (ITermData & ITerm)[] = await term.terms.get();
      let item: (ITermData & ITerm)[] = await term2.filter(i => {
        let indexLength = i.PathOfTerm.split(";").length;
        return i.PathOfTerm.split(";").length === indexLength;
      });

      return (await Promise.all<ISPTermObject>(item.map(async (t: any): Promise<ISPTermObject> => {
        return await this.projectTermAsync(t, siteUrl);
      })));


    }

    // Default empty array in case of any missing data
    return (new Promise<Array<ISPTermObject>>((resolve, reject) => {
      resolve(new Array<ISPTermObject>());
    }));
  }

  /**
   * @function
   * Projects a Term object into an object of type ISPTermObject, including child terms
   * @param guid
   */
  private async projectTermAsync(term: any, siteUrl): Promise<ISPTermObject> {

    return ({
      identity: term['_ObjectIdentity_'] !== undefined ? term['_ObjectIdentity_'] : "",
      isAvailableForTagging: term['IsAvailableForTagging'] !== undefined ? term['IsAvailableForTagging'] : false,
      guid: term['Id'] !== undefined ? this.cleanGuid(term['Id']) : "",
      name: term['Name'] !== undefined ? term['Name'] : "",
      customSortOrder: term['CustomSortOrder'] !== undefined ? term['CustomSortOrder'] : "",
      terms: await this.getChildTermsAsync(term, siteUrl),
      localCustomProperties: term['LocalCustomProperties'] !== undefined ? term['LocalCustomProperties'] : null,
    });
  }

  /**
   * @function
   * Clean the Guid from the Web Service response
   * @param guid
   */
  private cleanGuid(guid: string): string {
    if (guid !== undefined)
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    else
      return '';
  }
}
