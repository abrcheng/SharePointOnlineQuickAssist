import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
export default class RestoreItemsQAGrid extends React.Component<IRestoreItems>
{   
    public state = {items:this.props.items};
    private columns: IColumn[] = [       
        {
            key: 'Path',
            name: 'Path',
            fieldName: 'path',
            minWidth: 200,
            maxWidth: 330,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.Path}</span>;
            },
            isPadded: true,
          },
          {
            key: 'DeletedDate',
            name: 'DeletedDate',
            fieldName: 'deletedDate',
            minWidth: 100,
            maxWidth: 140,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.DeletedDate}</span>;
            },
            isPadded: true,
          },
          {
                key: 'DeletedByEmail',
                name: 'DeletedByEmail',
                fieldName: 'deletedByEmail',
                minWidth: 120,
                maxWidth: 180,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByEmail}</span>;
                },
                isPadded: true,
            },
            {
                key: 'DeletedByName',
                name: 'DeletedByName',
                fieldName: 'deletedByName',
                minWidth: 100,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByName}</span>;
                },
                isPadded: true,
            }          
      ];

    public render():React.ReactElement<IRestoreItems>
    {
        this.state.items = this.props.items;
        return <div>
             {this.state.items && this.state.items.length >0 ?
              <DetailsList
              items={this.state.items}            
              columns={this.columns}
              selectionMode={SelectionMode.none}            
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}            
            />: "No data"}            
        </div>;
    }
}