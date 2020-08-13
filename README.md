# hss sap-rpa
SAP Robotic Process Automation

# install
Use python 3.7 or pyautogui may have issues recognizing pixel color

```
pip install -r requirements.txt
touch src/orders.txt
```
# usage

- cnf: create confirmation file (open SAP export file)
- fixwbs: fix SN WBS mapping table (open SAP export file)
- inbox: SAP inbox tools (has issues)
- dostuff: main SAP-RPA script
  - remove: [CO02] remove added MATLCONS operations
  - manual: [CO02] add operations and components
  - add: [CO02] add operations and components
  - check: [CO02] check if operations added by winshuttle or manual
  - unconfirm: [CO13] unconfirm 0444 operations
  - unconfirm_part: [CO13] helpUnConfirmPart
  - delete: [CO02] set deletion flag
