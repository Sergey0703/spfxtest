import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import styles from './SpfxcentresWebPart.module.scss';

export interface ISpfxcentresWebPartProps {}

export interface ICentresItem {
  Id: number;
  Title: string;
  Respite: boolean;
}

export default class SpfxcentresWebPart extends BaseClientSideWebPart<ISpfxcentresWebPartProps> {

  private sp = spfi().using(SPFx(this.context));

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spfxcentres}">
        <div id="centresListContainer">
          <h2>Centres List</h2>
          <table id="centresTable" class="${styles.table}">
            <thead>
              <tr>
                <th>ID</th>
                <th>Title</th>
                <th>Respite</th>
              </tr>
            </thead>
            <tbody id="centresTableBody">
            </tbody>
          </table>
        </div>
      </div>`;

    this._fetchAndRenderListItems();
  }

  private async _fetchAndRenderListItems(): Promise<void> {
    try {
      const items: ICentresItem[] = await this.sp.web.lists.getByTitle('Centres').items.select('Id', 'Title', 'Respite')();

      let html: string = '';
      items.forEach((item: ICentresItem) => {
        html += `
          <tr>
            <td>${item.Id}</td>
            <td>${item.Title}</td>
            <td>${item.Respite ? 'Yes' : 'No'}</td>
          </tr>`;
      });

      const listContainer: Element | null = this.domElement.querySelector('#centresTableBody');
      if (listContainer) {
        listContainer.innerHTML = html;
      }
    } catch (error) {
      console.error('Error fetching list items:', error);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
