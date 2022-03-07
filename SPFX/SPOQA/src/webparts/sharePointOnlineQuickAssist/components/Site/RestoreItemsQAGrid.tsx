import { IRestoreItem,IRestoreItems } from "./IRestoreItem";
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as React from 'react';

export default class RestoreItemsQAGrid extends React.Component<IRestoreItems>
{   
    private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      let isSortedDescending = column.isSortedDescending;
      // If we've sorted this column, flip it.
        if (column.isSorted) {
          isSortedDescending = !isSortedDescending;
        }
      const { columns } = this.state;
      let { items } = this.state;
        
    // Sort the items.
    items = _copyAndSort(items, column.fieldName!, isSortedDescending);

    // Reset the items and columns to match the state.
    this.setState({
      items: items,
      columns: columns.map(col => {
        col.isSorted = col.key === column.key;
        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }
        return col;
      }),
    });
    }

    private columns: IColumn[] = [       
        {
            key: 'Path',
            name: 'Path',
            fieldName: 'Path',
            minWidth: 200,
            maxWidth: 330,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.Path}</span>;
            },            
            isPadded: true,
            onColumnClick: this._onColumnClick
          },
          {
            key: 'DeletedDate',
            name: 'DeletedDate',
            fieldName: 'DeletedDate',
            minWidth: 100,
            maxWidth: 140,
            isResizable: true,
            isCollapsible: true,
            data: 'string',          
            onRender: (item: IRestoreItem) => {
              return <span>{item.DeletedDate}</span>;
            },
            isPadded: true,
            onColumnClick: this._onColumnClick
          },
          {
                key: 'DeletedByEmail',
                name: 'DeletedByEmail',
                fieldName: 'DeletedByEmail',
                minWidth: 120,
                maxWidth: 180,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByEmail}</span>;
                },
                isPadded: true,
                onColumnClick: this._onColumnClick
            },
            {
                key: 'DeletedByName',
                name: 'DeletedByName',
                fieldName: 'DeletedByName',
                minWidth: 100,
                maxWidth: 120,
                isResizable: true,
                isCollapsible: true,
                data: 'string',          
                onRender: (item: IRestoreItem) => {
                return <span>{item.DeletedByName}</span>;
                },
                isPadded: true,
                onColumnClick: this._onColumnClick
            }          
      ];
     
    public state = {
        items:this.props.items,
        columns:this.columns};

    public render():React.ReactElement<IRestoreItems>
    {
        // this.state.items = this.props.items;
        const { items, columns } = this.state;
        return <div>
             {this.state.items && this.state.items.length >0 ?
              <DetailsList
              items={items}            
              columns={columns}
              selectionMode={SelectionMode.none}            
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}            
            />: "No data"}            
        </div>;
    }  
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
