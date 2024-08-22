import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './SpfxcentresWebPart.module.scss';

export interface ISpfxcentresWebPartProps {}

export interface ICentresItem {
  Id: number;
  Title: string;
  Respite: boolean;
}

export default class SpfxcentresWebPart extends BaseClientSideWebPart<ISpfxcentresWebPartProps> {

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
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Centres')/items?$select=Id,Title,Respite`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      
      if (response.ok) {
        const data = await response.json();
        const items: ICentresItem[] = data.value;

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
      } else {
        console.error(`Error fetching list items: ${response.statusText}`);
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
