import pandas as pd
import numpy as np
from collections import namedtuple

sales_order = '9082402'
prod_order = 167405406

col_names = ['level', 'material_number', 'material_desc', 'material_qty', 'material_unit', 'document_number', 'part_name']

cahu_data = [[0, 'CUSTOM-AHU', 'CUSTOM AIR HANDLING UNIT', 1, 'EA', np.nan, np.nan],]
cahu_df = pd.DataFrame(cahu_data, columns=col_names)

bom_df = pd.read_excel('9082402-0020 BOM.xlsx', 
                       names=['level', 'material_number', 'material_desc', 'material_qty', 'material_unit', 'document_number', 'part_name'])

# Combine the dataframes, resetting the indexes
bom_df = pd.concat([cahu_df, bom_df], ignore_index=True)


new_bom_list = []
# Iterate through doc_number and part_name columns to ensure they are in proper form
# Store all this in a new bom
for tup in bom_df.itertuples(index=True):
    curr_level = tup.level
    curr_material_number = tup.material_number
    curr_material_desc = tup.material_desc
    curr_material_qty = tup.material_qty
    curr_material_unit = tup.material_unit
    curr_document_number = tup.document_number
    curr_part_name = tup.part_name
    new_document_number = np.nan
    new_part_name = np.nan
    
    # check if curr_document_number has an appropriate entry, passing the information
    # into new_document_number if it is
    try: 
        new_document_number = int(curr_document_number)
    except ValueError:
        pass
    
    # Check if curr_part_name has an appropriate entry, passing the information
    # into new_part_name if it is
    try:
        if curr_part_name[:len(sales_order)] == sales_order:
            new_part_name = curr_part_name
    except (TypeError, IndexError):
        pass
    
    # Only append this entry if there is an entry in material_number
    if str(curr_material_number) != 'nan':
        new_row = [curr_level, 
                    curr_material_number, 
                    curr_material_desc, 
                    curr_material_qty, 
                    curr_material_unit, 
                    new_document_number, 
                    new_part_name]
        new_bom_list.append(new_row)

new_bom_df = pd.DataFrame(new_bom_list, columns=col_names)
    

new_bom_parent_list = []
parent_tuple = namedtuple('parent_tuple', ['index', 'level'])
parent_stack = [parent_tuple(index=0, level=0)]      # (List index, level)
# Now that we have removed rows, create parent list
for tup in new_bom_df.itertuples(index=True):
    curr_index = tup.Index
    curr_level = tup.level
    curr_material_number = tup.material_number
    curr_material_desc = tup.material_desc
    curr_material_qty = tup.material_qty
    curr_material_unit = tup.material_unit
    curr_document_number = tup.document_number
    curr_part_name = tup.part_name
    new_document_number = np.nan
    new_part_name = np.nan
    curr_parent_index = np.nan
    
    # Find this entries parent, if not the first entry in the list (the root)
    if curr_index > 0:
        # Peek the top of the list. Pop off entries until the level is exactly
        # 1 less than curr_level. At that point, we can mark the parent for this
        # entry, and append ourselves onto the end of the list
        top_parent = parent_stack[len(parent_stack) -1]
        desired_parent_level = curr_level - 1
        while top_parent.level != desired_parent_level:
            # pop this one off the stack
            parent_stack.pop()
            top_parent = parent_stack[len(parent_stack) -1]
        # Now that we are here, we know that the top parent is the correct level
        curr_parent_index = top_parent.index
        # Add ourselves to the stack
        parent_stack.append(parent_tuple(index=curr_index, level=curr_level))
        
    # Append this data
    new_parent_row = [curr_parent_index, curr_index]
    new_bom_parent_list.append(new_parent_row)
        
# Create a dataframe
bom_parent_df = pd.DataFrame(new_bom_parent_list, columns=['parent_bom_index', 'child_bom_index'])        
            

# Format dataframes for input
bom_parent_df.insert(0, 'job_prod_order', prod_order)

new_bom_df.insert(0, 'job_prod_order', prod_order)
new_bom_df = new_bom_df.rename_axis('bom_index').reset_index()


# Export to csv
bom_parent_df.to_csv('test_parent_relationships.csv', sep='|', index=False)
            
new_bom_df.to_csv('test_bom_output.csv', sep='|',
                  columns=['job_prod_order', 
                           'bom_index', 
                           'level', 
                           'material_number', 
                           'material_qty', 
                           'material_unit', 
                           'document_number', 
                           'part_name'],
                  index=False)
            
            
            
            
            
            
