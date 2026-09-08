package excel

import "embed"

//go:embed web/*
var webFS embed.FS

// AddinID is the Office add-in GUID (stable so WEF catalogs stay valid).
const AddinID = "5c1a1e15-0000-4000-a000-c01a1e150001"
