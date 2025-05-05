# Running the Refactored Code

## Current Testing/Validation

```powershell
python process_shipping_list_shim.py <packing_list_file> <policy_file> [output_dir]
```
This uses our shim that imports from the shipping_processor package but still calls the original implementation.

## Verify Compatibility

```powershell
python validation_program/run_validation.py --packing-list <packing_list_file> --policy-file <policy_file>
```
The validator imports the original process_shipping_list.py directly, so this validates the original code works.

## Switch to Refactored Code When Ready

```powershell
# Backup original
ren process_shipping_list.py process_shipping_list_original.py
# Put shim in place
ren process_shipping_list_shim.py process_shipping_list.py
```

## Development/Phased Migration

- Modify `shipping_processor/main.py` to incrementally replace parts of the process with refactored modules
- Test each change with validation program
- When all functionality is moved, remove the call to the original implementation and implement the function fully

The refactoring is designed to be incremental so you can migrate one function at a time while maintaining compatibility with validation. 