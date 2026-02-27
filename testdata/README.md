# Test Data

## Start.mdb

Source: `../MeKo/IPOffice-VBA/mdb-files/Start.mdb` (3.6 MB)
VBA Project Name: "Start"

### Expected VBA Modules (15 total)

| Module Name              | Type      |
|--------------------------|-----------|
| Form_Module              | DocClass  |
| Form_Module_Übersicht    | DocClass  |
| Form_Übersicht           | DocClass  |
| Form_Verweise            | DocClass  |
| Inidatei                 | Module    |
| Menueaufbau              | Module    |
| Modul1                   | Module    |
| SQL                      | Module    |
| IPAdresse2               | Module    |
| Form_Startschirm         | DocClass  |
| mod_api_Functions        | Module    |
| Update                   | Module    |
| Form_Mitteilung          | DocClass  |
| Verweise                 | Module    |
| ModAutostart             | Module    |

### VBA Storage Tree

```
MSysAccessStorage_ROOT (Id=1)
└── VBA (Id=8)
    └── VBAProject (Id=16)
        ├── PROJECT (Id=4847)         — plain-text module listing
        ├── PROJECTwm (Id=4846)       — unicode module name map
        └── VBA (Id=17)
            ├── _VBA_PROJECT (Id=4844)
            ├── dir (Id=4845)         — compressed module-stream mapping
            └── 15 module streams (Id=4815..4843, random names)
```

### Validation Criteria

- Extracted module count must equal 15
- Each `.bas` / `.cls` file should start with `Attribute VB_Name`
- Known keywords to verify: `DoCmd`, `MsgBox`, `Sub `, `Function `, `Dim `
- Standard modules (.bas): Inidatei, Menueaufbau, Modul1, SQL, IPAdresse2, mod_api_Functions, Update, Verweise, ModAutostart
- DocClass modules (.cls): Form_Module, Form_Module_Übersicht, Form_Übersicht, Form_Verweise, Form_Startschirm, Form_Mitteilung
