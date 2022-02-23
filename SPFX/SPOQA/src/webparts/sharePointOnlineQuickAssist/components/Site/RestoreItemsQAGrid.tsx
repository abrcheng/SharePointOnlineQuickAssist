import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';
export default class RestoreItemsQAGrid extends React.Component<IRestoreItems>
{   
    public state = {items:this.props.items};
    private columns: IColumn[] = [       
        {
            key: 'LeafName',
            name: 'FileName',
            fieldName: 'leafName',
            minWidth: 70,
            maxWidth: 90,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.LeafName}</span>;
            },
            isPadded: true,
          },      
          {
            key: 'DirName',
            name: 'Folder',
            fieldName: 'dirName',
            minWidth: 120,
            maxWidth: 200,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.DirName}</span>;
            },
            isPadded: true,
          },
          {
            key: 'DeletedDate',
            name: 'DeletedDate',
            fieldName: 'deletedDate',
            minWidth: 60,
            maxWidth: 80,
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
                minWidth: 70,
                maxWidth: 90,
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
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByEmail}</span>;
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