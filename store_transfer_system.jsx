import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from '@supabase/supabase-js';
import {
  Upload, Package, History, LayoutDashboard, Printer, AlertTriangle,
  CheckCircle2, Search, ArrowRightLeft, X, Trash2, ChevronRight,
  FileSpreadsheet, PlusCircle, Building2, Download, ListPlus, Pencil
  , ScanLine
} from "lucide-react";

const PRODUCTS_SEED = [
  {
    "sku": "RETAIL-031",
    "name": "ALLO 10K PODS",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-030",
    "name": "ALLO 10K PODS WITH BATTERY DEVICE",
    "category": "Vape",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-025",
    "name": "ALLO 1600",
    "category": "Vape",
    "unitPrice": 24.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-001",
    "name": "ALLO 800",
    "category": "Vape",
    "unitPrice": 19.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-001",
    "name": "ALLO 50k",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-034",
    "name": "ALLO 2500",
    "category": "Vape",
    "unitPrice": 29.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-023",
    "name": "ALPHA FB 80K",
    "category": "Vape",
    "unitPrice": 48.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-025",
    "name": "BAZOOKA 90K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-054",
    "name": "BLUE RAVEN / FRESH 60ML JUICE",
    "category": "Vape",
    "unitPrice": 38.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-033",
    "name": "BOOST ULTRA PRO G2 BATTERY",
    "category": "Vape",
    "unitPrice": 14.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-016",
    "name": "BREEZE 60K",
    "category": "Vape",
    "unitPrice": 49.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-002",
    "name": "BREEZE PRO 2000",
    "category": "Vape",
    "unitPrice": 29.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-009",
    "name": "CROWNBAR ALFAKHER 30K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-021",
    "name": "CROWNBAR ALFAKHER 75K",
    "category": "Vape",
    "unitPrice": 49.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-038",
    "name": "DRIP'N 16K",
    "category": "Vape",
    "unitPrice": 29.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-017",
    "name": "DRIPN 28K / DRIPN 30K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-055",
    "name": "DRIP'N 60ML JUICE",
    "category": "Vape",
    "unitPrice": 41.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-037",
    "name": "DRIP'N 70K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-006",
    "name": "DRIP'N 8ML, 16K",
    "category": "Vape",
    "unitPrice": 29.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-027",
    "name": "DRIP'N DAILY 100K",
    "category": "Vape",
    "unitPrice": 48.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-019",
    "name": "DRIP'N EVO 63K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-035",
    "name": "DRIP'N LEVEL X PODS BOOST G2 25K",
    "category": "Vape",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-037",
    "name": "DRIP'N LEVEL-X G2 ULTRA PODS 50K",
    "category": "Vape",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-046",
    "name": "E-JUICE / SALT BY BRAND NAME 30ML",
    "category": "Vape",
    "unitPrice": 34.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-047",
    "name": "E-JUICE / SALT BY BRAND NAME 60ML",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-031",
    "name": "ELFBAR 1800",
    "category": "Vape",
    "unitPrice": 18.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-031",
    "name": "ELFBAR 10K / DRIPN 10K",
    "category": "Vape",
    "unitPrice": 37.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-032",
    "name": "ELFBAR 20K",
    "category": "Vape",
    "unitPrice": 41.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-030",
    "name": "ELFBAR 70K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-022",
    "name": "ELFBAR BC PRO 80K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-022",
    "name": "ENVI APEX 2500",
    "category": "Vape",
    "unitPrice": 25.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-052",
    "name": "ENVY / OHMCITY 60ML JUICE",
    "category": "Vape",
    "unitPrice": 41.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-027",
    "name": "FB 50K BEAST MODE",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-028",
    "name": "FB LEVEL X 25K BOOST G2 PODS",
    "category": "Vape",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-029",
    "name": "FB LEVEL X 50K G2 ULTRA PODS",
    "category": "Vape",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-013",
    "name": "FB TWELVE MONKEYS 50K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-038",
    "name": "FB UNLEASHED LEVEL-X G2 ULTRA PODS 50K",
    "category": "Vape",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-057",
    "name": "FLAVOUR BEAST 120ML JUICE",
    "category": "Vape",
    "unitPrice": 51.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-026",
    "name": "FLAVOUR BEAST 18K MAX",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-056",
    "name": "FLAVOUR BEAST 30ML JUICE",
    "category": "Vape",
    "unitPrice": 34.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-056",
    "name": "FLAVOUR BEAST 60ML JUICE",
    "category": "Vape",
    "unitPrice": 51.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-036",
    "name": "FLAVOUR BEAST LEVEL X G2 ULTRA PODS 50K",
    "category": "Vape",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-032",
    "name": "FLAVOUR BEAST LEVEL X PODS BOOST G2 25K",
    "category": "Vape",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-014",
    "name": "FLAVOUR BEAST MAX 3 60K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-015",
    "name": "FLAVOUR BEAST MAX3 UNLEASHED 60K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-034",
    "name": "FLAVOUR BEAST UNLEASHED LEVEL X PODS BOOST G2 25K",
    "category": "Vape",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-007",
    "name": "FLAVOUR BEAST, BEAST MODE MAX 18K",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-005",
    "name": "FOG 15K",
    "category": "Vape",
    "unitPrice": 34.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-019",
    "name": "GEEK BAR 25K PULSE X / JNR SHISHA 30K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-018",
    "name": "GEEK BAR 9K PULSE",
    "category": "Vape",
    "unitPrice": 36.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-029",
    "name": "INSTABAR 120K",
    "category": "Vape",
    "unitPrice": 46.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-053",
    "name": "JAZZY 60ML JUICE",
    "category": "Vape",
    "unitPrice": 41.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900597564",
    "name": "KIT IQOS ILUMA I Breeze Blue",
    "category": "Vape",
    "unitPrice": 65.59,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900597595",
    "name": "KIT IQOS ILUMA I Midnight Black",
    "category": "Vape",
    "unitPrice": 65.59,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900597618",
    "name": "KIT IQOS ILUMA I ONE Breeze Blue",
    "category": "Vape",
    "unitPrice": 49.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900597649",
    "name": "KIT IQOS ILUMA I ONE Midnight Black",
    "category": "Vape",
    "unitPrice": 49.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-010",
    "name": "KRAZE LUNA 42K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-012",
    "name": "KRAZE MEGA 48K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-048",
    "name": "LEMON DROP 60ML JUICE",
    "category": "Vape",
    "unitPrice": 41.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-040",
    "name": "LOOP 3 BATTERY ONLY",
    "category": "Vape",
    "unitPrice": 4.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-042",
    "name": "LOOP MAX BATTERY ONLY",
    "category": "Vape",
    "unitPrice": 9.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-051",
    "name": "NITRO 60ML JUICE",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-028",
    "name": "ORBITO AI 120K",
    "category": "Vape",
    "unitPrice": 46.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-026",
    "name": "OXBAR 100K",
    "category": "Vape",
    "unitPrice": 49.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-045",
    "name": "OXBAR 90K PODS",
    "category": "Vape",
    "unitPrice": 36.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-024",
    "name": "OXBAR 90K STARTER KIT",
    "category": "Vape",
    "unitPrice": 48.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-043",
    "name": "OXBAR SVOPP PODS 60K",
    "category": "Vape",
    "unitPrice": 36.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-008",
    "name": "OXBAR SWOPP DEVICE",
    "category": "Vape",
    "unitPrice": 5.0,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-011",
    "name": "OXBAR TRI FUSION 45K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-010",
    "name": "RIFBAR 40K",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900602305",
    "name": "Sequoia Oak Birch Flavour Bundle",
    "category": "Vape",
    "unitPrice": 26.98,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-018",
    "name": "STLTH 60K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-004",
    "name": "STLTH 8K PRO",
    "category": "Vape",
    "unitPrice": 35.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-001",
    "name": "STLTH ECO 1600",
    "category": "Vape",
    "unitPrice": 18.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-004",
    "name": "STLTH ECO BOX 5K",
    "category": "Vape",
    "unitPrice": 27.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-008",
    "name": "STLTH ECO XL UP TO 20K",
    "category": "Vape",
    "unitPrice": 36.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-006",
    "name": "STLTH LOOP 25K PODS",
    "category": "Vape",
    "unitPrice": 32.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-039",
    "name": "STLTH LOOP PODS 25K",
    "category": "Vape",
    "unitPrice": 32.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-044",
    "name": "STLTH LOOP PODS 70K",
    "category": "Vape",
    "unitPrice": 37.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-003",
    "name": "STLTH TITAN MAX 50K/60K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-041",
    "name": "STLTH X ELFBAR LOOP PODS 50K",
    "category": "Vape",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RCPT-039",
    "name": "STLTH X GEEKBAR 80K",
    "category": "Vape",
    "unitPrice": 48.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900602176",
    "name": "TEREA ELM S50 PRI 180 SLI",
    "category": "Vape",
    "unitPrice": 68.5,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-049",
    "name": "TWELVE MONKEYS 60ML JUICE",
    "category": "Vape",
    "unitPrice": 51.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900399182",
    "name": "VEEV NOW Blue Mint 1.8% 18ml MNT",
    "category": "Vape",
    "unitPrice": 138.21,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900399199",
    "name": "VEEV NOW Spearmint 1.8% 18ml MNT",
    "category": "Vape",
    "unitPrice": 138.21,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VAPE-0061900399526",
    "name": "VEEV ONE Blue Mint 1.8% MNT OPK 2",
    "category": "Vape",
    "unitPrice": 59.08,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-017",
    "name": "VICE 60K",
    "category": "Vape",
    "unitPrice": 42.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-020",
    "name": "VICE 70K",
    "category": "Vape",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10200720",
    "name": "Vuse GO 5K 2.0 Banana Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10240435",
    "name": "Vuse GO 5K 2.0 Blueberry Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10240528",
    "name": "Vuse GO 5K 2.0 Clear 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10201885",
    "name": "Vuse GO 5K 2.0 Peach Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10240369",
    "name": "Vuse GO 5K 2.0 Smooth Tobacco 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10200762",
    "name": "Vuse GO 5K 2.0 Strawberry Kiwi 20mg 6x1",
    "category": "Vape",
    "unitPrice": 36.19,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10217920",
    "name": "Vuse GO 8K 2.0 Blueberry Raspberry 20mg 6x1",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10232440",
    "name": "Vuse GO 8K 2.0 Grape Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10217910",
    "name": "Vuse GO 8K 2.0 Green Apple 20mg 6x1",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10200785",
    "name": "Vuse GO 8K 2.0 Mint Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10200786",
    "name": "Vuse GO 8K 2.0 Spearmint Ice 20mg 6x1",
    "category": "Vape",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10228018",
    "name": "Vuse Pod 20mg Blueberry 5x2",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10228044",
    "name": "Vuse Pod 20mg Blueberry Raspberry 5x2",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10228065",
    "name": "Vuse Pod 20mg Peach 5x2",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10228037",
    "name": "Vuse Pod 20mg Rich Tobacco 5x2",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "VUSE-10228055",
    "name": "Vuse Pod 20mg Strawberry 5x2",
    "category": "Vape",
    "unitPrice": 21.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "RETAIL-003",
    "name": "Breeze 2500 ZERO NIC (NO DISCOUNT)",
    "category": "Vape",
    "unitPrice": 27.99,
    "priceType": "Per Unit",
    "packQty": 1
  },
  {
    "sku": "CONV-110",
    "name": "5 Gum Rain 10ct",
    "category": "Convenience",
    "unitPrice": 1.799,
    "priceType": "Pack Price",
    "packQty": 10,
    "subcategory": "Candy/Gum",
    "packPrice": 17.99
  },
  {
    "sku": "CONV-225",
    "name": "5 Gum Sour Strawberry 10ct",
    "category": "Convenience",
    "unitPrice": 17.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 17.99
  },
  {
    "sku": "CONV-029",
    "name": "7up Tropical Zero 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-271",
    "name": "Airhead Sour Blue Blast 36ct",
    "category": "Convenience",
    "unitPrice": 11.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Candy/Gum",
    "packPrice": 35.97
  },
  {
    "sku": "CONV-172",
    "name": "Airheads Bite Xtreme 18ct",
    "category": "Convenience",
    "unitPrice": 1.6661,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-109",
    "name": "Airheads Sour Watermelon Punch 36ct",
    "category": "Convenience",
    "unitPrice": 0.2358,
    "priceType": "Pack Price",
    "packQty": 36,
    "subcategory": "Candy/Gum",
    "packPrice": 8.49
  },
  {
    "sku": "CONV-163",
    "name": "Alani Cherry Slush 12x355ml",
    "category": "Convenience",
    "unitPrice": 2.2475,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-292",
    "name": "Alani Dream Float 12x355ml",
    "category": "Convenience",
    "unitPrice": 26.97,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-164",
    "name": "Alani Hawaiian Shaved Ice 12x355ml",
    "category": "Convenience",
    "unitPrice": 2.2475,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-215",
    "name": "Alani Lime Slush 12x355ml",
    "category": "Convenience",
    "unitPrice": 2.4992,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-145",
    "name": "Alani Pink Slush 12x355ml",
    "category": "Convenience",
    "unitPrice": 2.2475,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-144",
    "name": "Alani Winter Wonderland 12x355ml",
    "category": "Convenience",
    "unitPrice": 4.495,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 53.94
  },
  {
    "sku": "CONV-128",
    "name": "Almond Joy King Size 18x91g",
    "category": "Convenience",
    "unitPrice": 2.9439,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 52.99
  },
  {
    "sku": "CONV-034",
    "name": "Aloe Vera Original 20x500ml",
    "category": "Convenience",
    "unitPrice": 1.8995,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Beverages",
    "packPrice": 37.99
  },
  {
    "sku": "CONV-253",
    "name": "Arctic Ice Ice Cubes 1pk",
    "category": "Convenience",
    "unitPrice": 1.15,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Other",
    "packPrice": 9.2
  },
  {
    "sku": "CONV-004",
    "name": "Arizona Apple Lime Rickey 24x680ml",
    "category": "Convenience",
    "unitPrice": 1.1246,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-216",
    "name": "Arizona Frost Chillzicle 24x680ml USA",
    "category": "Convenience",
    "unitPrice": 0.9996,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 23.99
  },
  {
    "sku": "CONV-019",
    "name": "Arizona Green Tea 24x680ml CAD",
    "category": "Convenience",
    "unitPrice": 1.1871,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 28.49
  },
  {
    "sku": "CONV-022",
    "name": "Arizona Rizzler Berry 24x680ml USA",
    "category": "Convenience",
    "unitPrice": 1.1246,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-283",
    "name": "Arizona Tropical Chillzicle 24x680ml USA",
    "category": "Convenience",
    "unitPrice": 26.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-267",
    "name": "Bic Lighter Plain 8 50ct",
    "category": "Convenience",
    "unitPrice": 47.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Smoking Accessories",
    "packPrice": 95.98
  },
  {
    "sku": "CONV-049",
    "name": "Bic Lighter Plain Small 50ct",
    "category": "Convenience",
    "unitPrice": 0.9598,
    "priceType": "Pack Price",
    "packQty": 50,
    "subcategory": "Smoking Accessories",
    "packPrice": 47.99
  },
  {
    "sku": "CONV-178",
    "name": "Big League Gumballs Sour 3oz",
    "category": "Convenience",
    "unitPrice": 1.99,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 17.91
  },
  {
    "sku": "CONV-148",
    "name": "Blistex Lip Balm Ultra 4.25g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Health/Personal",
    "packPrice": 6.98
  },
  {
    "sku": "CONV-295",
    "name": "Bounty 12x91g",
    "category": "Convenience",
    "unitPrice": 32.49,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 32.49
  },
  {
    "sku": "CONV-167",
    "name": "Bugles Chili Cheese 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Chips/Snacks",
    "packPrice": 23.92
  },
  {
    "sku": "CONV-078",
    "name": "Bugles Cinnamon Toast Crunch 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 8.97
  },
  {
    "sku": "CONV-076",
    "name": "Bugles Hidden Valley Ranch 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 8.97
  },
  {
    "sku": "CONV-228",
    "name": "Bugles Hidden Valley Ranch 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Chips/Snacks",
    "packPrice": 17.94
  },
  {
    "sku": "CONV-229",
    "name": "Bugles Nacho Cheese 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 11.96
  },
  {
    "sku": "CONV-198",
    "name": "Bugles Original 3oz",
    "category": "Convenience",
    "unitPrice": 2.69,
    "priceType": "Pack Price",
    "packQty": 36,
    "subcategory": "Chips/Snacks",
    "packPrice": 96.84
  },
  {
    "sku": "CONV-281",
    "name": "Bugles Original 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 11,
    "subcategory": "Chips/Snacks",
    "packPrice": 32.89
  },
  {
    "sku": "CONV-077",
    "name": "Bugles Tabasco 85g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Chips/Snacks",
    "packPrice": 17.94
  },
  {
    "sku": "CONV-189",
    "name": "Cadbury Mini Puds 73g",
    "category": "Convenience",
    "unitPrice": 4.29,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 38.61
  },
  {
    "sku": "CONV-037",
    "name": "Canada Dry Cranberry 591ml",
    "category": "Convenience",
    "unitPrice": 1.89,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 45.36
  },
  {
    "sku": "CONV-239",
    "name": "Canada Dry Ginger Ale 1L V2 12ct",
    "category": "Convenience",
    "unitPrice": 1.69,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 20.28
  },
  {
    "sku": "CONV-245",
    "name": "Canada Dry Ginger Ale 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Beverages",
    "packPrice": 18.32
  },
  {
    "sku": "CONV-160",
    "name": "Canada Dry Ginger Ale Peach Mango Zero Cans 12oz 12ct",
    "category": "Convenience",
    "unitPrice": 2.2475,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-206",
    "name": "Canada Dry Gingerale 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 9.16
  },
  {
    "sku": "CONV-038",
    "name": "Canada Dry Peach 591ml",
    "category": "Convenience",
    "unitPrice": 1.89,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 45.36
  },
  {
    "sku": "CONV-036",
    "name": "Canada Dry Vanilla 591ml",
    "category": "Convenience",
    "unitPrice": 1.7246,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 41.39
  },
  {
    "sku": "CONV-301",
    "name": "CD Ale 24x500ml",
    "category": "Convenience",
    "unitPrice": 29.89,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 29.89
  },
  {
    "sku": "CONV-073",
    "name": "Cheetos Cheddar Jalapeno 54g",
    "category": "Convenience",
    "unitPrice": 0.93,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 2.79
  },
  {
    "sku": "CONV-052",
    "name": "Cheetos Puffs 26/75g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Chips/Snacks",
    "packPrice": 8.85
  },
  {
    "sku": "CONV-251",
    "name": "Christie Oreo Bubble Tea 263g",
    "category": "Convenience",
    "unitPrice": 4.19,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Candy/Gum",
    "packPrice": 29.33
  },
  {
    "sku": "CONV-252",
    "name": "Christie Oreo Marvel 303g",
    "category": "Convenience",
    "unitPrice": 5.79,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 46.32
  },
  {
    "sku": "CONV-246",
    "name": "Coke 2L Single BTL Diet",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 9.16
  },
  {
    "sku": "CONV-298",
    "name": "Coke 32x355ml",
    "category": "Convenience",
    "unitPrice": 16.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 16.99
  },
  {
    "sku": "CONV-025",
    "name": "Coke All Kind 24x500ml",
    "category": "Convenience",
    "unitPrice": 1.8746,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-142",
    "name": "Coke Diet 500ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.39,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 33.36
  },
  {
    "sku": "CONV-294",
    "name": "Coke Diet 500ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.39,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 33.36
  },
  {
    "sku": "CONV-026",
    "name": "Coke Reg 24x355ml",
    "category": "Convenience",
    "unitPrice": 2.9146,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 69.95
  },
  {
    "sku": "CONV-005",
    "name": "Coke Reg 500ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.39,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 33.36
  },
  {
    "sku": "CONV-204",
    "name": "Coke Reg Bottle 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Beverages",
    "packPrice": 13.74
  },
  {
    "sku": "CONV-242",
    "name": "Coke Reg BTL 1L",
    "category": "Convenience",
    "unitPrice": 1.69,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 20.28
  },
  {
    "sku": "CONV-241",
    "name": "Coke Reg BTL 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 27.48
  },
  {
    "sku": "CONV-001",
    "name": "Coke Vanilla 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-212",
    "name": "Coke Vanilla Zero 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-299",
    "name": "Coke Zero 24ct",
    "category": "Convenience",
    "unitPrice": 16.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 16.99
  },
  {
    "sku": "CONV-202",
    "name": "Coke Zero Bottle 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 9.16
  },
  {
    "sku": "CONV-244",
    "name": "Coke Zero BTL 2L",
    "category": "Convenience",
    "unitPrice": 2.29,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 9.16
  },
  {
    "sku": "CONV-282",
    "name": "Core Power Chocolate 42g 12x414ml",
    "category": "Convenience",
    "unitPrice": 50.99,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Beverages",
    "packPrice": 254.95
  },
  {
    "sku": "CONV-285",
    "name": "Core Power Elite Vanilla 42g 12x414ml",
    "category": "Convenience",
    "unitPrice": 54.99,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Beverages",
    "packPrice": 274.95
  },
  {
    "sku": "CONV-009",
    "name": "Crush 80's Electric Blue Razz 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-064",
    "name": "Doritos Jalapeno Cheddar Family Size 235g",
    "category": "Convenience",
    "unitPrice": 3.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 7.98
  },
  {
    "sku": "CONV-055",
    "name": "Doritos Jalapeno Fro 32/72g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Chips/Snacks",
    "packPrice": 12.39
  },
  {
    "sku": "CONV-054",
    "name": "Doritos Nacho 32/72g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Chips/Snacks",
    "packPrice": 8.85
  },
  {
    "sku": "CONV-056",
    "name": "Doritos Zesty Cheese 32/72g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Chips/Snacks",
    "packPrice": 8.85
  },
  {
    "sku": "CONV-293",
    "name": "Dr Pepper Cherry Zero 355ml",
    "category": "Convenience",
    "unitPrice": 12.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Beverages",
    "packPrice": 25.98
  },
  {
    "sku": "CONV-263",
    "name": "Dunkin Brownie Batter Creme-filled Egg 1oz 24-Pack",
    "category": "Convenience",
    "unitPrice": 33.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 33.99
  },
  {
    "sku": "CONV-196",
    "name": "Durex Performax Mutual Climax Condoms 3pk",
    "category": "Convenience",
    "unitPrice": 18.6633,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Health/Personal",
    "packPrice": 55.99
  },
  {
    "sku": "CONV-118",
    "name": "E.Frutti Gummi Cupcake 60ct",
    "category": "Convenience",
    "unitPrice": 13.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 13.99
  },
  {
    "sku": "CONV-157",
    "name": "E.Frutti Mini Burger 600g 60ct",
    "category": "Convenience",
    "unitPrice": 13.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 13.99
  },
  {
    "sku": "CONV-117",
    "name": "E.Frutti Mini Burger 60ct",
    "category": "Convenience",
    "unitPrice": 13.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 13.99
  },
  {
    "sku": "CONV-045",
    "name": "Eagle Torch Pen Neon 12ct",
    "category": "Convenience",
    "unitPrice": 4.9992,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Smoking Accessories",
    "packPrice": 59.99
  },
  {
    "sku": "CONV-080",
    "name": "E-frutti Sweet and Sour Apple 100g",
    "category": "Convenience",
    "unitPrice": 1.69,
    "priceType": "Pack Price",
    "packQty": 11,
    "subcategory": "Candy/Gum",
    "packPrice": 18.59
  },
  {
    "sku": "CONV-033",
    "name": "Evian 12x1.5L",
    "category": "Convenience",
    "unitPrice": 2.4992,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-249",
    "name": "Evian Spring Water 50cl/500ml 24ct",
    "category": "Convenience",
    "unitPrice": 20.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 20.99
  },
  {
    "sku": "CONV-272",
    "name": "Excel Bubblemint 12x12ct",
    "category": "Convenience",
    "unitPrice": 13.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 13.99
  },
  {
    "sku": "CONV-169",
    "name": "Excel Spearmint 12x12",
    "category": "Convenience",
    "unitPrice": 1.1658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 13.99
  },
  {
    "sku": "CONV-238",
    "name": "Exclusive Candy Shock Rocks Popping Watermelon 24ct",
    "category": "Convenience",
    "unitPrice": 14.59,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 14.59
  },
  {
    "sku": "CONV-146",
    "name": "Fanta Lemon 24x330ml",
    "category": "Convenience",
    "unitPrice": 1.2496,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-211",
    "name": "Fanta Lemon 24x330ml",
    "category": "Convenience",
    "unitPrice": 1.2496,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-011",
    "name": "Fanta Peach 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-213",
    "name": "Fanta Pineapple & Grapefruit 24x330ml",
    "category": "Convenience",
    "unitPrice": 1.2496,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-010",
    "name": "Fanta Pineapple 12x355ml",
    "category": "Convenience",
    "unitPrice": 0.7492,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 8.99
  },
  {
    "sku": "CONV-231",
    "name": "Fini Sour Rollers Blueraspberry 40ct",
    "category": "Convenience",
    "unitPrice": 35.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 35.99
  },
  {
    "sku": "CONV-174",
    "name": "Finn Tornadoes Smooth Raspberry Pencils 140g UK",
    "category": "Convenience",
    "unitPrice": 2.09,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 50.16
  },
  {
    "sku": "CONV-047",
    "name": "Fisherman Friend Lemon 24ct",
    "category": "Convenience",
    "unitPrice": 1.8121,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 43.49
  },
  {
    "sku": "CONV-048",
    "name": "Fisherman Friend Original 8ct",
    "category": "Convenience",
    "unitPrice": 1.8113,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 14.49
  },
  {
    "sku": "CONV-203",
    "name": "Fresca Bottle 2L",
    "category": "Convenience",
    "unitPrice": 2.39,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Beverages",
    "packPrice": 4.78
  },
  {
    "sku": "CONV-074",
    "name": "Fritos Hoops BBQ 57g",
    "category": "Convenience",
    "unitPrice": 0.93,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 3.72
  },
  {
    "sku": "CONV-190",
    "name": "Fun Dip Candy Cherry Orange Grape 40.5g 36ct",
    "category": "Convenience",
    "unitPrice": 1.2497,
    "priceType": "Pack Price",
    "packQty": 36,
    "subcategory": "Candy/Gum",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-134",
    "name": "Garbage Bag 40ct 26x32 67lb",
    "category": "Convenience",
    "unitPrice": 0.0848,
    "priceType": "Pack Price",
    "packQty": 40,
    "subcategory": "Household",
    "packPrice": 3.39
  },
  {
    "sku": "CONV-028",
    "name": "Gatorade All Kind 24x710ml",
    "category": "Convenience",
    "unitPrice": 1.9163,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-218",
    "name": "Gatorade Fruit Punch 710ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.89,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 7.56
  },
  {
    "sku": "CONV-207",
    "name": "Gatorade Zero Glacier Freeze 591ml Can",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Beverages",
    "packPrice": 6.54
  },
  {
    "sku": "CONV-014",
    "name": "Ghost Blue Raspberry 12x473ml 180mg CAD",
    "category": "Convenience",
    "unitPrice": 2.9992,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 35.99
  },
  {
    "sku": "CONV-042",
    "name": "Gill Slide 4pk #563",
    "category": "Convenience",
    "unitPrice": 2.9975,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Smoking Accessories",
    "packPrice": 11.99
  },
  {
    "sku": "CONV-232",
    "name": "Grenade Protein Bar Cookie Dough 12x60g",
    "category": "Convenience",
    "unitPrice": 35.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Snacks",
    "packPrice": 35.99
  },
  {
    "sku": "CONV-149",
    "name": "Gushers Super Sour 8ct",
    "category": "Convenience",
    "unitPrice": 2.4987,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 19.99
  },
  {
    "sku": "CONV-273",
    "name": "Halls Cool Mint Can 20ct",
    "category": "Convenience",
    "unitPrice": 24.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-111",
    "name": "Halls Honey Lemon 20ct",
    "category": "Convenience",
    "unitPrice": 1.2495,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-095",
    "name": "Haribo Berry Hearts 3.1oz 12pk",
    "category": "Convenience",
    "unitPrice": 1.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 19.99
  },
  {
    "sku": "CONV-092",
    "name": "Haribo Bitter Lemon and Friends Gummies 160g 20pk",
    "category": "Convenience",
    "unitPrice": 3.2995,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Candy/Gum",
    "packPrice": 65.99
  },
  {
    "sku": "CONV-223",
    "name": "Haribo Fizzy Cola 142g",
    "category": "Convenience",
    "unitPrice": 2.19,
    "priceType": "Pack Price",
    "packQty": 5,
    "subcategory": "Candy/Gum",
    "packPrice": 10.95
  },
  {
    "sku": "CONV-280",
    "name": "Haribo Fizzy Cola 142g",
    "category": "Convenience",
    "unitPrice": 2.19,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 19.71
  },
  {
    "sku": "CONV-116",
    "name": "Haribo Gummies Tangfastics 175g",
    "category": "Convenience",
    "unitPrice": 2.25,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 27
  },
  {
    "sku": "CONV-098",
    "name": "Haribo Happy Cherries Gummi Candy 12pk",
    "category": "Convenience",
    "unitPrice": 2.3325,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 27.99
  },
  {
    "sku": "CONV-079",
    "name": "Haribo Happy Cola 142g",
    "category": "Convenience",
    "unitPrice": 2.19,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 13.14
  },
  {
    "sku": "CONV-099",
    "name": "Haribo Harry Potter Ron 6.3oz 10pk",
    "category": "Convenience",
    "unitPrice": 2.999,
    "priceType": "Pack Price",
    "packQty": 10,
    "subcategory": "Candy/Gum",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-256",
    "name": "Haribo Heart Throbs 140g UK 12-Pack",
    "category": "Convenience",
    "unitPrice": 35.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 35.99
  },
  {
    "sku": "CONV-093",
    "name": "Haribo Raspberries Gummi Candy 12pk",
    "category": "Convenience",
    "unitPrice": 2.3325,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 27.99
  },
  {
    "sku": "CONV-091",
    "name": "Haribo Soda Twist 176g 12pk",
    "category": "Convenience",
    "unitPrice": 2.8,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 33.6
  },
  {
    "sku": "CONV-090",
    "name": "Haribo Sour Soda Zing 4.5oz 12pk",
    "category": "Convenience",
    "unitPrice": 2.2492,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-097",
    "name": "Haribo Super Wummis 205g 14pk",
    "category": "Convenience",
    "unitPrice": 3.285,
    "priceType": "Pack Price",
    "packQty": 14,
    "subcategory": "Candy/Gum",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-096",
    "name": "Haribo Sweet and Sour Hearts 9oz 8pk",
    "category": "Convenience",
    "unitPrice": 4.7488,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 37.99
  },
  {
    "sku": "CONV-094",
    "name": "Haribo Watermelon Gummi Candy 12pk",
    "category": "Convenience",
    "unitPrice": 2.3325,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 27.99
  },
  {
    "sku": "CONV-255",
    "name": "Hershey's Caramel & Sea Salt 95g 14-Pack",
    "category": "Convenience",
    "unitPrice": 49.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 49.99
  },
  {
    "sku": "CONV-159",
    "name": "Hershey's Ice Breakers Sour Watermelon Green Apple Tangerine 6x43g",
    "category": "Convenience",
    "unitPrice": 12.19,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 12.19
  },
  {
    "sku": "CONV-084",
    "name": "Hershey's Nougat Bar 18x51g",
    "category": "Convenience",
    "unitPrice": 1.1661,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 20.99
  },
  {
    "sku": "CONV-185",
    "name": "Hi-Chew Chewy Candy Acai 15ct 1.73oz",
    "category": "Convenience",
    "unitPrice": 1.5327,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Candy/Gum",
    "packPrice": 22.99
  },
  {
    "sku": "CONV-186",
    "name": "Hi-Chew Chewy Candy Banana 15ct 1.76oz",
    "category": "Convenience",
    "unitPrice": 1.5327,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Candy/Gum",
    "packPrice": 22.99
  },
  {
    "sku": "CONV-081",
    "name": "Hi-Chew Stick Blue Raspberry 15x50g",
    "category": "Convenience",
    "unitPrice": 1.5327,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Candy/Gum",
    "packPrice": 22.99
  },
  {
    "sku": "CONV-152",
    "name": "Hi-Chew Watermelon 15x50g",
    "category": "Convenience",
    "unitPrice": 1.5327,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Candy/Gum",
    "packPrice": 22.99
  },
  {
    "sku": "CONV-154",
    "name": "Ice Breaker Sparkling Pineapple Mango 8ct",
    "category": "Convenience",
    "unitPrice": 3.1237,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-155",
    "name": "Ice Breakers Cherry Limeade 8ct",
    "category": "Convenience",
    "unitPrice": 3.1237,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-153",
    "name": "Ice Breakers Mint Spearmint 8ct",
    "category": "Convenience",
    "unitPrice": 3.1237,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-021",
    "name": "Ice River Green Water 12x500ml",
    "category": "Convenience",
    "unitPrice": 0.2492,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 2.99
  },
  {
    "sku": "CONV-205",
    "name": "Ice River Green Water 500ml 12ct",
    "category": "Convenience",
    "unitPrice": 2.49,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Beverages",
    "packPrice": 4.98
  },
  {
    "sku": "CONV-030",
    "name": "Iced Coffee Oreo 12x473ml",
    "category": "Convenience",
    "unitPrice": 3.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 43.99
  },
  {
    "sku": "CONV-031",
    "name": "Iced Coffee Reese's 12x473ml",
    "category": "Convenience",
    "unitPrice": 3.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 43.99
  },
  {
    "sku": "CONV-277",
    "name": "Jolly Rancher Hard Candy Original 198g",
    "category": "Convenience",
    "unitPrice": 3.69,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Candy/Gum",
    "packPrice": 25.83
  },
  {
    "sku": "CONV-171",
    "name": "Jumbo Sour Suckers 1.2kg",
    "category": "Convenience",
    "unitPrice": 9.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 19.98
  },
  {
    "sku": "CONV-230",
    "name": "Jumbo Sour Suckers 1.2kg",
    "category": "Convenience",
    "unitPrice": 9.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 19.98
  },
  {
    "sku": "CONV-147",
    "name": "Kinder Bueno 20x43g",
    "category": "Convenience",
    "unitPrice": 1.2495,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Candy/Gum",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-180",
    "name": "Kinder Chocolate 8 pcs",
    "category": "Convenience",
    "unitPrice": 2.6162,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 20.93
  },
  {
    "sku": "CONV-264",
    "name": "Kit Kat Cookie Dough 99g 15-Pack",
    "category": "Convenience",
    "unitPrice": 65.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 65.99
  },
  {
    "sku": "CONV-265",
    "name": "Kit Kat Mega Chunky Salted Caramel 68g 24-Pack",
    "category": "Convenience",
    "unitPrice": 71.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 71.99
  },
  {
    "sku": "CONV-133",
    "name": "KitKat Regular 48x45g",
    "category": "Convenience",
    "unitPrice": 1.1248,
    "priceType": "Pack Price",
    "packQty": 48,
    "subcategory": "Candy/Gum",
    "packPrice": 53.99
  },
  {
    "sku": "CONV-173",
    "name": "Kool Aid U1 12x473ml",
    "category": "Convenience",
    "unitPrice": 4.1658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 49.99
  },
  {
    "sku": "CONV-071",
    "name": "Kurkure Chilli Chatka Chips 115g",
    "category": "Convenience",
    "unitPrice": 1.59,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 3.18
  },
  {
    "sku": "CONV-057",
    "name": "Kurkure Masala 30/115g",
    "category": "Convenience",
    "unitPrice": 1.46,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Chips/Snacks",
    "packPrice": 1.46
  },
  {
    "sku": "CONV-070",
    "name": "Large Chips Lays BBQ 60g",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 4.36
  },
  {
    "sku": "CONV-069",
    "name": "Large Chips Miss Vickies Original 55g",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 2.18
  },
  {
    "sku": "CONV-059",
    "name": "Lay's BBQ 27/66g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 5.31
  },
  {
    "sku": "CONV-058",
    "name": "Lay's Classic 27/66g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Chips/Snacks",
    "packPrice": 12.39
  },
  {
    "sku": "CONV-050",
    "name": "Lays Cream and Onion 18/200g",
    "category": "Convenience",
    "unitPrice": 4.54,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 18.16
  },
  {
    "sku": "CONV-051",
    "name": "Lays Salt and Vinegar 27/66g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 5.31
  },
  {
    "sku": "CONV-257",
    "name": "Lay's Stax Spicy Lobster 100g Thailand 16-Pack",
    "category": "Convenience",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Chips/Snacks",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-259",
    "name": "Lay's Stax Wagyu Beef & Truffle 100g Thailand 16-Pack",
    "category": "Convenience",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Chips/Snacks",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-260",
    "name": "Lay's Wavy Texas Tenderloin Steak 90g Vietnam 40-Pack",
    "category": "Convenience",
    "unitPrice": 89.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Chips/Snacks",
    "packPrice": 89.99
  },
  {
    "sku": "CONV-151",
    "name": "Love Rose 6in 24ct",
    "category": "Convenience",
    "unitPrice": 1.2492,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Smoking Accessories",
    "packPrice": 29.98
  },
  {
    "sku": "CONV-199",
    "name": "M&M Chocolate Bark Heart 99g",
    "category": "Convenience",
    "unitPrice": 6.99,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 41.94
  },
  {
    "sku": "CONV-191",
    "name": "M&M Honey Roasted Peanut 24ct 1.74oz",
    "category": "Convenience",
    "unitPrice": 0.5967,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 14.32
  },
  {
    "sku": "CONV-193",
    "name": "M&M Pop'd Caramel Pouch Freeze Dried Candy 5.5oz",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 26.91
  },
  {
    "sku": "CONV-300",
    "name": "Maison Glass 24ct",
    "category": "Convenience",
    "unitPrice": 23.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 23.99
  },
  {
    "sku": "CONV-268",
    "name": "Mars Regular 48x52g",
    "category": "Convenience",
    "unitPrice": 59.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 59.99
  },
  {
    "sku": "CONV-150",
    "name": "Maven Chrome Torch 15ct",
    "category": "Convenience",
    "unitPrice": 3.9993,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Smoking Accessories",
    "packPrice": 59.99
  },
  {
    "sku": "CONV-044",
    "name": "Metal Pipe With Screen 12ct",
    "category": "Convenience",
    "unitPrice": 1.49,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Smoking Accessories",
    "packPrice": 17.88
  },
  {
    "sku": "CONV-270",
    "name": "Mike & Ike Ice Cream Truck 12ct",
    "category": "Convenience",
    "unitPrice": 1.69,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 20.28
  },
  {
    "sku": "CONV-122",
    "name": "Milka Biscoff Lotus 90g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 42.96
  },
  {
    "sku": "CONV-120",
    "name": "Milka Extra Cacao 90g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 25,
    "subcategory": "Candy/Gum",
    "packPrice": 44.75
  },
  {
    "sku": "CONV-192",
    "name": "Milka Happy Cow 90g",
    "category": "Convenience",
    "unitPrice": 8.69,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Candy/Gum",
    "packPrice": 60.83
  },
  {
    "sku": "CONV-184",
    "name": "Milka Happy Cow Bubbly White 100g",
    "category": "Convenience",
    "unitPrice": 1.99,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Candy/Gum",
    "packPrice": 7.96
  },
  {
    "sku": "CONV-181",
    "name": "Milka LU Chocolate Sandwich 87g",
    "category": "Convenience",
    "unitPrice": 1.99,
    "priceType": "Pack Price",
    "packQty": 10,
    "subcategory": "Candy/Gum",
    "packPrice": 19.9
  },
  {
    "sku": "CONV-123",
    "name": "Milka Milkinis 87.5g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Candy/Gum",
    "packPrice": 35.8
  },
  {
    "sku": "CONV-124",
    "name": "Milka Oreo 100g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 22,
    "subcategory": "Candy/Gum",
    "packPrice": 39.38
  },
  {
    "sku": "CONV-121",
    "name": "Milka Oreo Brownie 100g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 26,
    "subcategory": "Candy/Gum",
    "packPrice": 46.54
  },
  {
    "sku": "CONV-182",
    "name": "Milka Raspberry Cream Chocolate 100g",
    "category": "Convenience",
    "unitPrice": 1.99,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 23.88
  },
  {
    "sku": "CONV-125",
    "name": "Milka White Chocolate 90g",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 42.96
  },
  {
    "sku": "CONV-130",
    "name": "Milky Way All Caramel 24x43g",
    "category": "Convenience",
    "unitPrice": 5.9988,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 143.97
  },
  {
    "sku": "CONV-068",
    "name": "Miss Vickies Jalapeno 40g",
    "category": "Convenience",
    "unitPrice": 0.93,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 1.86
  },
  {
    "sku": "CONV-053",
    "name": "Miss Vickies Sea Salt and Malt 36/59g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 7.08
  },
  {
    "sku": "CONV-161",
    "name": "Monster Absolutely Zero 12ct 473ml Can",
    "category": "Convenience",
    "unitPrice": 2.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-013",
    "name": "Monster All Kind 12x473ml",
    "category": "Convenience",
    "unitPrice": 2.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-288",
    "name": "Monster All Kind 12x473ml",
    "category": "Convenience",
    "unitPrice": 24.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 24.99
  },
  {
    "sku": "CONV-289",
    "name": "Monster Variety Pack 24x473ml",
    "category": "Convenience",
    "unitPrice": 45.97,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 45.97
  },
  {
    "sku": "CONV-297",
    "name": "Monster Zero 12ct",
    "category": "Convenience",
    "unitPrice": 44.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-127",
    "name": "Mounds King Size 18x99g",
    "category": "Convenience",
    "unitPrice": 5.8878,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 105.98
  },
  {
    "sku": "CONV-012",
    "name": "Mount Dew Voltage 355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 25.98
  },
  {
    "sku": "CONV-002",
    "name": "Mountain Dew Baja Blast Citrus 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-003",
    "name": "Mountain Dew Live Wire 355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 25.98
  },
  {
    "sku": "CONV-243",
    "name": "Naya Natural Spring Water 1L 12ct",
    "category": "Convenience",
    "unitPrice": 6.29,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 6.29
  },
  {
    "sku": "CONV-008",
    "name": "Naya Water 12x1L",
    "category": "Convenience",
    "unitPrice": 0.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 7.99
  },
  {
    "sku": "CONV-170",
    "name": "Nelson 2% Partly Skimmed Milk 473ml",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Dairy/Food",
    "packPrice": 1.09
  },
  {
    "sku": "CONV-114",
    "name": "Nerds Gummy Clusters 141g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Candy/Gum",
    "packPrice": 8.97
  },
  {
    "sku": "CONV-276",
    "name": "Nerds Gummy Clusters 141g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 13,
    "subcategory": "Candy/Gum",
    "packPrice": 38.87
  },
  {
    "sku": "CONV-275",
    "name": "Nerds Gummy Clusters Very Berry 142g",
    "category": "Convenience",
    "unitPrice": 2.99,
    "priceType": "Pack Price",
    "packQty": 15,
    "subcategory": "Candy/Gum",
    "packPrice": 44.85
  },
  {
    "sku": "CONV-106",
    "name": "Nerds Rainbow Pink Gummy Clusters 3oz 12pk",
    "category": "Convenience",
    "unitPrice": 3.2492,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 38.99
  },
  {
    "sku": "CONV-279",
    "name": "Nibo Lighter Refill 162g",
    "category": "Convenience",
    "unitPrice": 2.49,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Smoking Accessories",
    "packPrice": 29.88
  },
  {
    "sku": "CONV-126",
    "name": "Nutella Biscuits 304g",
    "category": "Convenience",
    "unitPrice": 6.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 13.98
  },
  {
    "sku": "CONV-043",
    "name": "Oil Burner Color 4in 4ct",
    "category": "Convenience",
    "unitPrice": 1.2475,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Smoking Accessories",
    "packPrice": 4.99
  },
  {
    "sku": "CONV-035",
    "name": "Orange Chronic 16oz",
    "category": "Convenience",
    "unitPrice": 7.49,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 89.88
  },
  {
    "sku": "CONV-219",
    "name": "Oreo Cakesters 8x3.03oz",
    "category": "Convenience",
    "unitPrice": 20.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 41.98
  },
  {
    "sku": "CONV-129",
    "name": "Payday Chocolatey King Size 18x96g",
    "category": "Convenience",
    "unitPrice": 2.9439,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 52.99
  },
  {
    "sku": "CONV-248",
    "name": "Peace Tea Caddy Shack 695ml 12ct",
    "category": "Convenience",
    "unitPrice": 13.29,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 13.29
  },
  {
    "sku": "CONV-250",
    "name": "Peace Tea Razzleberry 695ml 12ct",
    "category": "Convenience",
    "unitPrice": 13.29,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 13.29
  },
  {
    "sku": "CONV-247",
    "name": "Peace Tea Sno-Berry 695ml 12ct",
    "category": "Convenience",
    "unitPrice": 13.29,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 13.29
  },
  {
    "sku": "CONV-087",
    "name": "Peelerz Gummy Grape 170g",
    "category": "Convenience",
    "unitPrice": 3.79,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 22.74
  },
  {
    "sku": "CONV-088",
    "name": "Peelerz Gummy Lychee 170g",
    "category": "Convenience",
    "unitPrice": 3.79,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 22.74
  },
  {
    "sku": "CONV-089",
    "name": "Peelerz Gummy Mango 170g",
    "category": "Convenience",
    "unitPrice": 3.79,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 22.74
  },
  {
    "sku": "CONV-086",
    "name": "Peelerz Gummy Orange 170g",
    "category": "Convenience",
    "unitPrice": 3.79,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 22.74
  },
  {
    "sku": "CONV-085",
    "name": "Peelerz Gummy Peach 170g",
    "category": "Convenience",
    "unitPrice": 3.79,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 22.74
  },
  {
    "sku": "CONV-007",
    "name": "Pepsi Maple Cola 591ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.49,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 35.76
  },
  {
    "sku": "CONV-027",
    "name": "Pepsi Real Sugar 12x355ml",
    "category": "Convenience",
    "unitPrice": 1.0825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 12.99
  },
  {
    "sku": "CONV-020",
    "name": "Pepsi Reg 591ml 1ct",
    "category": "Convenience",
    "unitPrice": 1.49,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 35.76
  },
  {
    "sku": "CONV-217",
    "name": "Pepsi Zero Cream Soda 24x330ml",
    "category": "Convenience",
    "unitPrice": 1.1246,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-006",
    "name": "Pepsi Zero Strawberry & Cream 24x330ml",
    "category": "Convenience",
    "unitPrice": 0.9996,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 23.99
  },
  {
    "sku": "CONV-240",
    "name": "Perrier Maison Lime Sparkling Water 500ml 24ct",
    "category": "Convenience",
    "unitPrice": 30.59,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 30.59
  },
  {
    "sku": "CONV-143",
    "name": "Powerade All Kind 12x710ml",
    "category": "Convenience",
    "unitPrice": 3.3317,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 39.98
  },
  {
    "sku": "CONV-278",
    "name": "Pringles Carnitas Taco USA 158g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 14,
    "subcategory": "Chips/Snacks",
    "packPrice": 48.86
  },
  {
    "sku": "CONV-254",
    "name": "Pringles Flame Grilled Steak 165g UK 19-Pack",
    "category": "Convenience",
    "unitPrice": 88.49,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Chips/Snacks",
    "packPrice": 88.49
  },
  {
    "sku": "CONV-075",
    "name": "Rap Snacks Chips All Kind 71g",
    "category": "Convenience",
    "unitPrice": 2.19,
    "priceType": "Pack Price",
    "packQty": 14,
    "subcategory": "Chips/Snacks",
    "packPrice": 30.66
  },
  {
    "sku": "CONV-220",
    "name": "Raw Black King Size 50ct",
    "category": "Convenience",
    "unitPrice": 45.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Smoking Accessories",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-221",
    "name": "Raw Classic King Size 50ct",
    "category": "Convenience",
    "unitPrice": 39.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Smoking Accessories",
    "packPrice": 39.99
  },
  {
    "sku": "CONV-040",
    "name": "Raw Cone King Size 32ct",
    "category": "Convenience",
    "unitPrice": 1.4372,
    "priceType": "Pack Price",
    "packQty": 32,
    "subcategory": "Smoking Accessories",
    "packPrice": 45.99
  },
  {
    "sku": "CONV-041",
    "name": "Raw Cone Tips 20ct",
    "category": "Convenience",
    "unitPrice": 0.9745,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Smoking Accessories",
    "packPrice": 19.49
  },
  {
    "sku": "CONV-119",
    "name": "Red Band Jumbo Sour Suckers 1.2kg 60ct",
    "category": "Convenience",
    "unitPrice": 9.49,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 9.49
  },
  {
    "sku": "CONV-296",
    "name": "Red Bull 473ml",
    "category": "Convenience",
    "unitPrice": 44.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-290",
    "name": "Red Bull Cherry Sakura 24x250ml",
    "category": "Convenience",
    "unitPrice": 69.99,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Beverages",
    "packPrice": 279.96
  },
  {
    "sku": "CONV-291",
    "name": "Red Bull Coconut Berry 24x250ml",
    "category": "Convenience",
    "unitPrice": 69.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 69.99
  },
  {
    "sku": "CONV-165",
    "name": "Redbull Coconut Berry 12x250ml",
    "category": "Convenience",
    "unitPrice": 5.8317,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 69.98
  },
  {
    "sku": "CONV-023",
    "name": "Redbull Peach 4pk 24x250ml",
    "category": "Convenience",
    "unitPrice": 2.0204,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 48.49
  },
  {
    "sku": "CONV-016",
    "name": "Redbull Pink 24x250ml",
    "category": "Convenience",
    "unitPrice": 2.9162,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 69.99
  },
  {
    "sku": "CONV-032",
    "name": "Redbull Reg 12x473ml Large",
    "category": "Convenience",
    "unitPrice": 3.5742,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 42.89
  },
  {
    "sku": "CONV-015",
    "name": "Redbull Small Zero 24x250ml",
    "category": "Convenience",
    "unitPrice": 1.8746,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-024",
    "name": "Redbull Winter Edition Fuji Apple 24x250ml",
    "category": "Convenience",
    "unitPrice": 2.9162,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 69.99
  },
  {
    "sku": "CONV-166",
    "name": "Redbull Winter Edition Fuji Apple 24x355ml",
    "category": "Convenience",
    "unitPrice": 11.2488,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 269.97
  },
  {
    "sku": "CONV-131",
    "name": "Reeses Big Cup Sugar Cookies King Size 16x73g",
    "category": "Convenience",
    "unitPrice": 4.8738,
    "priceType": "Pack Price",
    "packQty": 16,
    "subcategory": "Candy/Gum",
    "packPrice": 77.98
  },
  {
    "sku": "CONV-132",
    "name": "Reese's P.B. With Oreo Crumbs 24x39g",
    "category": "Convenience",
    "unitPrice": 1.8746,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-194",
    "name": "Ritter Sport Dark Chocolate 50% Cocoa 100g",
    "category": "Convenience",
    "unitPrice": 26.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 26.99
  },
  {
    "sku": "CONV-156",
    "name": "Ruffles BBQ Flaming Hot Family Size 190g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 10.47
  },
  {
    "sku": "CONV-063",
    "name": "Ruffles Sour Cream and Onion 60g",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 2.18
  },
  {
    "sku": "CONV-237",
    "name": "Samyang Carbonara Buldak Cup 80g 6ct",
    "category": "Convenience",
    "unitPrice": 9.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Snacks",
    "packPrice": 19.98
  },
  {
    "sku": "CONV-138",
    "name": "Samyang Cheese Chick Ramen Cup 6x70g",
    "category": "Convenience",
    "unitPrice": 3.4967,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Ramen/Food",
    "packPrice": 20.98
  },
  {
    "sku": "CONV-141",
    "name": "Samyang Chick Carbo Ramen Cup 6x70g",
    "category": "Convenience",
    "unitPrice": 4.995,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Ramen/Food",
    "packPrice": 29.97
  },
  {
    "sku": "CONV-236",
    "name": "Samyang Hot Chicken Ramen 75g 6ct",
    "category": "Convenience",
    "unitPrice": 9.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Snacks",
    "packPrice": 9.99
  },
  {
    "sku": "CONV-140",
    "name": "Samyang Spicy Chick Ramen Cup 6x70g",
    "category": "Convenience",
    "unitPrice": 1.7483,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Ramen/Food",
    "packPrice": 10.49
  },
  {
    "sku": "CONV-208",
    "name": "Sanpellegrino Water 250ml 24ct",
    "category": "Convenience",
    "unitPrice": 1.0829,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Beverages",
    "packPrice": 25.99
  },
  {
    "sku": "CONV-139",
    "name": "Shin Ramyun Black 6x101g",
    "category": "Convenience",
    "unitPrice": 5.6633,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Ramen/Food",
    "packPrice": 33.98
  },
  {
    "sku": "CONV-101",
    "name": "Skittles Desserts 125g 12pk",
    "category": "Convenience",
    "unitPrice": 3.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 43.99
  },
  {
    "sku": "CONV-100",
    "name": "Skittles Fruit Chewies 125g 12pk",
    "category": "Convenience",
    "unitPrice": 3.0833,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 37
  },
  {
    "sku": "CONV-262",
    "name": "Skittles Gummies Colourful Fruit 42.5g China 12-Pack",
    "category": "Convenience",
    "unitPrice": 20.49,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 20.49
  },
  {
    "sku": "CONV-115",
    "name": "Skittles Gummies Original 164g",
    "category": "Convenience",
    "unitPrice": 3.09,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Candy/Gum",
    "packPrice": 9.27
  },
  {
    "sku": "CONV-168",
    "name": "Skittles Gummies Sour 164g",
    "category": "Convenience",
    "unitPrice": 3.09,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Candy/Gum",
    "packPrice": 12.36
  },
  {
    "sku": "CONV-234",
    "name": "Skittles Gummies Sour 164g",
    "category": "Convenience",
    "unitPrice": 3.09,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Candy/Gum",
    "packPrice": 9.27
  },
  {
    "sku": "CONV-258",
    "name": "Skittles Gummies Yogurt Fruit 42.5g China 12-Pack",
    "category": "Convenience",
    "unitPrice": 20.49,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 20.49
  },
  {
    "sku": "CONV-287",
    "name": "Skittles Juice Grape 12x414ml",
    "category": "Convenience",
    "unitPrice": 29.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Beverages",
    "packPrice": 89.97
  },
  {
    "sku": "CONV-286",
    "name": "Skittles Juice Lime 12x414ml",
    "category": "Convenience",
    "unitPrice": 29.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Beverages",
    "packPrice": 29.99
  },
  {
    "sku": "CONV-284",
    "name": "Skittles Juice Orange 12x414ml",
    "category": "Convenience",
    "unitPrice": 29.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Beverages",
    "packPrice": 89.97
  },
  {
    "sku": "CONV-112",
    "name": "Skittles Pop'd Original Pouch 155g",
    "category": "Convenience",
    "unitPrice": 8.99,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Candy/Gum",
    "packPrice": 35.96
  },
  {
    "sku": "CONV-233",
    "name": "Skittles Pop'd Original Pouch 155g",
    "category": "Convenience",
    "unitPrice": 8.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 17.98
  },
  {
    "sku": "CONV-113",
    "name": "Skittles Pop'd Sour Pouch 155g",
    "category": "Convenience",
    "unitPrice": 8.99,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Candy/Gum",
    "packPrice": 35.96
  },
  {
    "sku": "CONV-227",
    "name": "Skittles Pop'd Sour Pouch 155g",
    "category": "Convenience",
    "unitPrice": 8.99,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Candy/Gum",
    "packPrice": 17.98
  },
  {
    "sku": "CONV-188",
    "name": "Skittles Sour Pop'd Peg 5.5oz",
    "category": "Convenience",
    "unitPrice": 8.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Candy/Gum",
    "packPrice": 26.97
  },
  {
    "sku": "CONV-274",
    "name": "Sky USB to iPhone 20ct",
    "category": "Convenience",
    "unitPrice": 19.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Accessories",
    "packPrice": 19.99
  },
  {
    "sku": "CONV-061",
    "name": "Smartfood White Cheddar 26/50g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 7.08
  },
  {
    "sku": "CONV-072",
    "name": "Smartfood White Cheddar Popcorn 45g",
    "category": "Convenience",
    "unitPrice": 1.09,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Chips/Snacks",
    "packPrice": 2.18
  },
  {
    "sku": "CONV-018",
    "name": "Snapple Kiwi Strawberry 12x473ml",
    "category": "Convenience",
    "unitPrice": 0.9992,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 11.99
  },
  {
    "sku": "CONV-197",
    "name": "Snickers Chocolate Bark Heart 99g",
    "category": "Convenience",
    "unitPrice": 6.99,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 41.94
  },
  {
    "sku": "CONV-183",
    "name": "Sour Patch Kids Bites Tube 12ct 1.8oz",
    "category": "Convenience",
    "unitPrice": 2.4158,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 28.99
  },
  {
    "sku": "CONV-175",
    "name": "Sour Patch Kids Chews Box 18ct 1.94oz",
    "category": "Convenience",
    "unitPrice": 2.2217,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 39.99
  },
  {
    "sku": "CONV-108",
    "name": "Sour Patch Kids Glowups Easter Eggs 8.37oz 12pk",
    "category": "Convenience",
    "unitPrice": 5.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 67.99
  },
  {
    "sku": "CONV-162",
    "name": "Sour Patch Kids Glowups Valentine 8.4oz 12 Pack",
    "category": "Convenience",
    "unitPrice": 5.6658,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 67.99
  },
  {
    "sku": "CONV-103",
    "name": "Sour Punch Filled Straws Cherry Limeade 5oz 12pk",
    "category": "Convenience",
    "unitPrice": 3.3333,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 40
  },
  {
    "sku": "CONV-105",
    "name": "Sour Punch Straws 3.2oz 24pk",
    "category": "Convenience",
    "unitPrice": 2.0829,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 49.99
  },
  {
    "sku": "CONV-104",
    "name": "Sour Punch Straws Blue Raspberry 2oz 24pk",
    "category": "Convenience",
    "unitPrice": 1.4579,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 34.99
  },
  {
    "sku": "CONV-224",
    "name": "Sour Strips Bites Blue Razz 180g",
    "category": "Convenience",
    "unitPrice": 5.2789,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 47.51
  },
  {
    "sku": "CONV-226",
    "name": "Sour Strips Bites Rainbow 180g",
    "category": "Convenience",
    "unitPrice": 5.29,
    "priceType": "Pack Price",
    "packQty": 7,
    "subcategory": "Candy/Gum",
    "packPrice": 37.03
  },
  {
    "sku": "CONV-067",
    "name": "Stacy's Pita Chips Simply Naked 38.3g",
    "category": "Convenience",
    "unitPrice": 0.99,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 2.97
  },
  {
    "sku": "CONV-062",
    "name": "Sun Chips Harvest Cheddar 27/70g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 5.31
  },
  {
    "sku": "CONV-176",
    "name": "Swedish Fish Chews Box 18ct 1.94oz",
    "category": "Convenience",
    "unitPrice": 2.2217,
    "priceType": "Pack Price",
    "packQty": 18,
    "subcategory": "Candy/Gum",
    "packPrice": 39.99
  },
  {
    "sku": "CONV-066",
    "name": "Takis Blue Heat Tortilla Chips 80g",
    "category": "Convenience",
    "unitPrice": 1.49,
    "priceType": "Pack Price",
    "packQty": 4,
    "subcategory": "Chips/Snacks",
    "packPrice": 5.96
  },
  {
    "sku": "CONV-107",
    "name": "Takis Chile Limon 3.25oz 20pk",
    "category": "Convenience",
    "unitPrice": 3.0,
    "priceType": "Pack Price",
    "packQty": 20,
    "subcategory": "Chips/Snacks",
    "packPrice": 60
  },
  {
    "sku": "CONV-065",
    "name": "Takis Fuego Extreme Tortilla Chips 80g",
    "category": "Convenience",
    "unitPrice": 1.49,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 4.47
  },
  {
    "sku": "CONV-060",
    "name": "Tostitos BSR Loaded Nacho 36/80g",
    "category": "Convenience",
    "unitPrice": 1.77,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 5.31
  },
  {
    "sku": "CONV-179",
    "name": "Trident Sour Patch Kids Blue Raspberry 14 pcs x 12",
    "category": "Convenience",
    "unitPrice": 1.2158,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Candy/Gum",
    "packPrice": 14.59
  },
  {
    "sku": "CONV-177",
    "name": "Trident Vibes Blue Raspberry 40pk 6ct",
    "category": "Convenience",
    "unitPrice": 6.6633,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 39.98
  },
  {
    "sku": "CONV-082",
    "name": "Trident White Peppermint 9x16ct",
    "category": "Convenience",
    "unitPrice": 1.7767,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 15.99
  },
  {
    "sku": "CONV-083",
    "name": "Trident White Winterfresh 9x16ct",
    "category": "Convenience",
    "unitPrice": 1.7767,
    "priceType": "Pack Price",
    "packQty": 9,
    "subcategory": "Candy/Gum",
    "packPrice": 15.99
  },
  {
    "sku": "CONV-102",
    "name": "Trolli Apple Mallow Filled Marshmallows 150g 8pk",
    "category": "Convenience",
    "unitPrice": 4.4988,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Candy/Gum",
    "packPrice": 35.99
  },
  {
    "sku": "CONV-158",
    "name": "Turtle Chips Choco Churros 160g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 3,
    "subcategory": "Chips/Snacks",
    "packPrice": 10.47
  },
  {
    "sku": "CONV-200",
    "name": "Turtle Chips K-Chicken 160g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Chips/Snacks",
    "packPrice": 27.92
  },
  {
    "sku": "CONV-201",
    "name": "Turtle Chips Mexican 160g",
    "category": "Convenience",
    "unitPrice": 3.49,
    "priceType": "Pack Price",
    "packQty": 8,
    "subcategory": "Chips/Snacks",
    "packPrice": 27.92
  },
  {
    "sku": "CONV-266",
    "name": "Twix Bits Share Size 2.83oz 12-Pack",
    "category": "Convenience",
    "unitPrice": 44.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 44.99
  },
  {
    "sku": "CONV-137",
    "name": "Vaseline Blueseal Original 50ml",
    "category": "Convenience",
    "unitPrice": 1.89,
    "priceType": "Pack Price",
    "packQty": 2,
    "subcategory": "Health/Personal",
    "packPrice": 3.78
  },
  {
    "sku": "CONV-136",
    "name": "Vaseline Jar Cocoa Butter 7g",
    "category": "Convenience",
    "unitPrice": 2.69,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Health/Personal",
    "packPrice": 2.69
  },
  {
    "sku": "CONV-135",
    "name": "Vaseline Jar Creme Brulee 7g",
    "category": "Convenience",
    "unitPrice": 2.69,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Health/Personal",
    "packPrice": 2.69
  },
  {
    "sku": "CONV-017",
    "name": "Vita Coconut Water 12x330ml",
    "category": "Convenience",
    "unitPrice": 1.5825,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 18.99
  },
  {
    "sku": "CONV-214",
    "name": "Vitamin Water All 12x591ml",
    "category": "Convenience",
    "unitPrice": 3.665,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 43.98
  },
  {
    "sku": "CONV-209",
    "name": "Vitamin Water Single Mega C Dragonfruit 591ml Can",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 21.48
  },
  {
    "sku": "CONV-210",
    "name": "Vitamin Water Single Zero Squeezed 591ml Can",
    "category": "Convenience",
    "unitPrice": 1.79,
    "priceType": "Pack Price",
    "packQty": 12,
    "subcategory": "Beverages",
    "packPrice": 21.48
  },
  {
    "sku": "CONV-235",
    "name": "Wang Spicy Topokki Carbonara 169g",
    "category": "Convenience",
    "unitPrice": 4.99,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Snacks",
    "packPrice": 29.94
  },
  {
    "sku": "CONV-195",
    "name": "Wonka Bottle Caps 50g 24ct",
    "category": "Convenience",
    "unitPrice": 0.5375,
    "priceType": "Pack Price",
    "packQty": 24,
    "subcategory": "Candy/Gum",
    "packPrice": 12.9
  },
  {
    "sku": "CONV-187",
    "name": "Wrigley's Excel Refreshers Strawberry 6x40 Pieces",
    "category": "Convenience",
    "unitPrice": 3.9983,
    "priceType": "Pack Price",
    "packQty": 6,
    "subcategory": "Candy/Gum",
    "packPrice": 23.99
  },
  {
    "sku": "CONV-222",
    "name": "X-lite Flip Lighter 30ct",
    "category": "Convenience",
    "unitPrice": 31.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Smoking Accessories",
    "packPrice": 31.99
  },
  {
    "sku": "CONV-261",
    "name": "Yorkie Raisin and Biscuit 44g UK 24-Pack",
    "category": "Convenience",
    "unitPrice": 48.0,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Candy/Gum",
    "packPrice": 48
  },
  {
    "sku": "CONV-046",
    "name": "Zengaz Countertop Lighters 48ct",
    "category": "Convenience",
    "unitPrice": 2.9165,
    "priceType": "Pack Price",
    "packQty": 48,
    "subcategory": "Smoking Accessories",
    "packPrice": 139.99
  },
  {
    "sku": "CONV-269",
    "name": "Zippo Fluid 4floz 12x118ml",
    "category": "Convenience",
    "unitPrice": 36.99,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Smoking Accessories",
    "packPrice": 36.99
  },
  {
    "sku": "CONV-039",
    "name": "Zippo Fluid Small 4oz",
    "category": "Convenience",
    "unitPrice": 4.18,
    "priceType": "Per Unit",
    "packQty": 1,
    "subcategory": "Smoking Accessories",
    "packPrice": 4.18
  },
  {
    "sku": "CIG-006190001122",
    "name": "Belmont KS S&S 20",
    "category": "Cigarettes",
    "unitPrice": 14.594,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-006190001139",
    "name": "Belmont KS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 18.2425,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-006190001146",
    "name": "Belmont RS 20",
    "category": "Cigarettes",
    "unitPrice": 14.594,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-006190001153",
    "name": "Belmont RS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 18.2425,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-0061900038128",
    "name": "Belmont Select King Size 20 (SS)",
    "category": "Cigarettes",
    "unitPrice": 14.594,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-0061900038135",
    "name": "Belmont Select KS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 18.2425,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-0061900038142",
    "name": "Belmont Select RS 20",
    "category": "Cigarettes",
    "unitPrice": 14.594,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-0061900038159",
    "name": "Belmont Select RS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 18.2425,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Belmont"
  },
  {
    "sku": "CIG-006190003997",
    "name": "Benson & Hedges Deluxe KS S&S 20",
    "category": "Cigarettes",
    "unitPrice": 15.116,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Benson & Hedges"
  },
  {
    "sku": "CIG-0061900010124",
    "name": "Benson & Hedges Unison KS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 18.895,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Benson & Hedges"
  },
  {
    "sku": "CIG-0061900200785",
    "name": "Benson and Hedges KS 25",
    "category": "Cigarettes",
    "unitPrice": 18.895,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Benson & Hedges"
  },
  {
    "sku": "CIGAR-0061900574053",
    "name": "Cigrill Sail Classic 100 FLP 1 Fat",
    "category": "Cigarettes",
    "unitPrice": 40.93,
    "priceType": "Per Unit",
    "packQty": "1",
    "brand": "Cigrill"
  },
  {
    "sku": "CIGAR-0061900574008",
    "name": "Cigrill Sail Classic 100 FLP 8 Fat",
    "category": "Cigarettes",
    "unitPrice": 15.6325,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Cigrill"
  },
  {
    "sku": "CIGAR-0061900573988",
    "name": "Cigrill Sail Highland 100 FLP 8 Fat",
    "category": "Cigarettes",
    "unitPrice": 15.6325,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Cigrill"
  },
  {
    "sku": "TOB-006190060237",
    "name": "Drum Premier (Blue) 50g",
    "category": "Cigarettes",
    "unitPrice": 167.82,
    "priceType": "Per Unit",
    "packQty": "50g",
    "brand": "Drum"
  },
  {
    "sku": "CIG-0061900506125",
    "name": "du Maurier Distinct Plus KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900506132",
    "name": "du Maurier Distinct Plus KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 18.565,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900505623",
    "name": "du Maurier Distinct RG 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900520039",
    "name": "du Maurier Mellow KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900520053",
    "name": "du Maurier Mellow RG 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900502913",
    "name": "du Maurier Signature KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900502936",
    "name": "du Maurier Signature RG 10x20",
    "category": "Cigarettes",
    "unitPrice": 14.852,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "du Maurier"
  },
  {
    "sku": "CIG-0061900561704",
    "name": "Marlboro Original KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 14.6487,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Marlboro"
  },
  {
    "sku": "CIG-0061900563647",
    "name": "Marlboro Smooth KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 14.6487,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Marlboro"
  },
  {
    "sku": "CIG-0061900504428",
    "name": "Matinee Mellow KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 22.0637,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Matinee"
  },
  {
    "sku": "CIG-0061900510580",
    "name": "Matinee Subtle KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 22.0637,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Matinee"
  },
  {
    "sku": "CIG-0061900007001",
    "name": "Next KS 20",
    "category": "Cigarettes",
    "unitPrice": 11.395,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900007056",
    "name": "Next KS 25",
    "category": "Cigarettes",
    "unitPrice": 14.2438,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900007230",
    "name": "Next P M Original KS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 14.2438,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900007223",
    "name": "Next P M Special KS S&S 25",
    "category": "Cigarettes",
    "unitPrice": 14.2438,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900007377",
    "name": "Next Smooth KS 25",
    "category": "Cigarettes",
    "unitPrice": 14.2438,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900007025",
    "name": "Next Xtra KS 20",
    "category": "Cigarettes",
    "unitPrice": 11.395,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Next"
  },
  {
    "sku": "CIG-0061900527773",
    "name": "Pall Mall Bold KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 14.2988,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-0061900562770",
    "name": "Pall Mall Full KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 14.2988,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-006190055341",
    "name": "Pall Mall Smooth Extra KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 11.439,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-006190052408",
    "name": "Pall Mall Smooth KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 11.439,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-006190052415",
    "name": "Pall Mall Smooth KS 8x25",
    "category": "Cigarettes",
    "unitPrice": 14.2988,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-0061900524785",
    "name": "Pall Mall Smooth RG 10x20",
    "category": "Cigarettes",
    "unitPrice": 11.439,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Pall Mall"
  },
  {
    "sku": "CIG-0061900552993",
    "name": "Player's Original RG 8x25",
    "category": "Cigarettes",
    "unitPrice": 20.2838,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Player's"
  },
  {
    "sku": "CIG-0061900521782",
    "name": "Player's RG 8x25",
    "category": "Cigarettes",
    "unitPrice": 20.2838,
    "priceType": "Per Unit",
    "packQty": 8,
    "brand": "Player's"
  },
  {
    "sku": "CIG-0061900563340",
    "name": "Vogue KS 10x20",
    "category": "Cigarettes",
    "unitPrice": 17.601,
    "priceType": "Per Unit",
    "packQty": 10,
    "brand": "Vogue"
  },
  {
    "sku": "BEER-006",
    "name": "Bud Light 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 66.96,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0001107",
    "depositPerUnit": 0,
    "casePrice": 66.96
  },
  {
    "sku": "BEER-007",
    "name": "Budweiser 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 66.96,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0905976",
    "depositPerUnit": 0,
    "casePrice": 66.96
  },
  {
    "sku": "BEER-008",
    "name": "Corona Extra 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 82.08,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0017853",
    "depositPerUnit": 0,
    "casePrice": 82.08
  },
  {
    "sku": "BEER-011",
    "name": "Cottage Springs Ontario Peach",
    "category": "Beer & Alcohol",
    "unitPrice": 79.2,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0553537",
    "depositPerUnit": 0,
    "casePrice": 79.2
  },
  {
    "sku": "BEER-AUTO-452",
    "name": "Henekiean",
    "category": "Beer & Alcohol",
    "unitPrice": 4.125,
    "priceType": "Per Unit (from case price)",
    "packQty": 24,
    "supplier": "LCBO",
    "lcbo": null,
    "depositPerUnit": 0,
    "casePrice": 99
  },
  {
    "sku": "BEER-001",
    "name": "Labatt Blue 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 2.3,
    "priceType": "Per Unit (from case price)",
    "packQty": 24,
    "supplier": "LCBO",
    "lcbo": "6696468",
    "depositPerUnit": 2.4,
    "casePrice": 55.2
  },
  {
    "sku": "BEER-002",
    "name": "Labatt Blue Ice 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 2.26,
    "priceType": "Per Unit (from case price)",
    "packQty": 24,
    "supplier": "LCBO",
    "lcbo": "0449520",
    "depositPerUnit": 2.4,
    "casePrice": 54.24
  },
  {
    "sku": "BEER-004",
    "name": "Laker Ice 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 2.26,
    "priceType": "Per Unit (from case price)",
    "packQty": 24,
    "supplier": "LCBO",
    "lcbo": "0142620",
    "depositPerUnit": 2.4,
    "casePrice": 54.24
  },
  {
    "sku": "BEER-005",
    "name": "Lowenbrau 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 55.2,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0397638",
    "depositPerUnit": 0,
    "casePrice": 55.2
  },
  {
    "sku": "BEER-013",
    "name": "Mikes Harder Lemonade",
    "category": "Beer & Alcohol",
    "unitPrice": 36.36,
    "priceType": "Per Unit (from case price)",
    "packQty": 2,
    "supplier": "LCBO",
    "lcbo": "0036697",
    "depositPerUnit": 0,
    "casePrice": 72.72
  },
  {
    "sku": "BEER-010",
    "name": "Nutrl Grape",
    "category": "Beer & Alcohol",
    "unitPrice": 64.08,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0031355",
    "depositPerUnit": 0,
    "casePrice": 64.08
  },
  {
    "sku": "BEER-003",
    "name": "Stella Artois 473ml",
    "category": "Beer & Alcohol",
    "unitPrice": 3.39,
    "priceType": "Per Unit (from case price)",
    "packQty": 24,
    "supplier": "LCBO",
    "lcbo": "0017820",
    "depositPerUnit": 2.4,
    "casePrice": 81.36
  },
  {
    "sku": "BEER-009",
    "name": "White Claw Hard Seltzer Black Cherry",
    "category": "Beer & Alcohol",
    "unitPrice": 75.84,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0014486",
    "depositPerUnit": 0,
    "casePrice": 75.84
  },
  {
    "sku": "BEER-012",
    "name": "White Claw Surge Blood Orange",
    "category": "Beer & Alcohol",
    "unitPrice": 79.2,
    "priceType": "Per Unit (from case price)",
    "packQty": 1,
    "supplier": "LCBO",
    "lcbo": "0031288",
    "depositPerUnit": 0,
    "casePrice": 79.2
  }
];

const DEFAULT_STORES = ["Dixie", "Vaughan"];
const HST_RATE = 0.13;

const CATEGORY_COLORS = {
  "Vape": { bg: "#eef2ff", text: "#4338ca", dot: "#6366f1" },
  "Convenience": { bg: "#fef3e2", text: "#b45309", dot: "#e8a33d" },
  "Cigarettes": { bg: "#fdf2f2", text: "#a5202a", dot: "#c0392b" },
  "Beer & Alcohol": { bg: "#effaf3", text: "#166534", dot: "#2f9e58" },
};

function normalize(s) {
  return String(s || "").trim().toUpperCase().replace(/\s+/g, " ");
}

function findHeaderKey(headers, candidates) {
  const norm = headers.map(h => ({ raw: h, n: normalize(h) }));
  for (const c of candidates) {
    const hit = norm.find(h => h.n === normalize(c));
    if (hit) return hit.raw;
  }
  for (const c of candidates) {
    const hit = norm.find(h => h.n.includes(normalize(c)));
    if (hit) return hit.raw;
  }
  return null;
}

function buildSkuIndex(products) {
  const map = {};
  products.forEach(p => {
    if (!map[p.sku]) map[p.sku] = [];
    map[p.sku].push(p);
  });
  return map;
}

function scoreNameMatch(query, name) {
  const q = normalize(query);
  const n = normalize(name);
  if (q === n) return 100;
  if (n.includes(q) || q.includes(n)) return 70;
  const qWords = new Set(q.split(" ").filter(Boolean));
  const nWords = new Set(n.split(" ").filter(Boolean));
  let overlap = 0;
  qWords.forEach(w => { if (nWords.has(w)) overlap++; });
  return overlap > 0 ? 30 + overlap * 5 : 0;
}

function matchProduct(row, skuIndex, products) {
  // Product name is the primary match key — it's what people actually remember.
  const nameRaw = row.productName ? String(row.productName).trim() : "";
  if (nameRaw) {
    let best = null, bestScore = 0;
    for (const p of products) {
      const score = scoreNameMatch(nameRaw, p.name);
      if (score > bestScore) { bestScore = score; best = p; }
    }
    if (best && bestScore >= 70) return { product: best, status: "matched", candidates: [best] };
    if (best && bestScore >= 30) return { product: best, status: "fuzzy", candidates: [best] };
  }
  // SKU is only used as a fallback when the name didn't give a confident match —
  // a mismatched or missing SKU should never block an otherwise-good name match.
  const skuRaw = row.sku ? String(row.sku).trim() : "";
  if (skuRaw && skuIndex[skuRaw]) {
    const matches = skuIndex[skuRaw];
    return { product: matches[0], status: matches.length > 1 ? "ambiguous" : "matched", candidates: matches };
  }
  return { product: null, status: "unmatched", candidates: [] };
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function money(n) {
  return "$" + (Number(n) || 0).toLocaleString("en-CA", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

async function loadTransfers() {
  try {
    const SUPA_URL = import.meta.env.VITE_SUPABASE_URL;
    const SUPA_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
    const supabase = SUPA_URL && SUPA_KEY ? createClient(SUPA_URL, SUPA_KEY) : null;
    if (supabase) {
      // load transfers and their items from Supabase
      const { data: tdata, error: terr } = await supabase.from('transfers').select('*').order('date', { ascending: true });
      if (terr) throw terr;
      if (!tdata || !tdata.length) return [];
      const ids = tdata.map(t => t.id);
      const { data: items, error: ierr } = await supabase.from('transfer_items').select('*').in('transfer_id', ids);
      if (ierr) throw ierr;
      const itemsById = {};
      (items || []).forEach(it => {
        itemsById[it.transfer_id] = itemsById[it.transfer_id] || [];
        itemsById[it.transfer_id].push({ sku: it.sku, name: it.name, category: it.category, qty: it.qty, unitPrice: it.unit_price ?? it.unitPrice, lineTotal: it.line_total ?? it.lineTotal, metadata: it.metadata });
      });
      return tdata.map(t => ({ id: t.id, date: t.date, fromStore: t.from_store || t.fromStore, toStore: t.to_store || t.toStore, notes: t.notes, subtotal: t.subtotal, hst: t.hst, total: t.total, createdAt: t.created_at, reconciledWith: t.reconciled_with, raw: t.raw, items: itemsById[t.id] || [] }));
    }
    if (window.storage?.get) {
      const res = await window.storage.get("transfers", true);
      return res ? JSON.parse(res.value) : [];
    }
    return JSON.parse(window.localStorage.getItem("store-transfer:transfers") || "[]");
  } catch (e) {
    console.error('loadTransfers error', e);
    return [];
  }
}

async function saveTransfers(transfers) {
  try {
    const SUPA_URL = import.meta.env.VITE_SUPABASE_URL;
    const SUPA_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
    const supabase = SUPA_URL && SUPA_KEY ? createClient(SUPA_URL, SUPA_KEY) : null;
    if (supabase) {
      // upsert transfers and transfer_items
      for (const t of transfers) {
        const up = {
          id: t.id,
          date: t.date,
          from_store: t.fromStore,
          to_store: t.toStore,
          notes: t.notes,
          subtotal: t.subtotal,
          hst: t.hst,
          total: t.total,
          created_at: t.createdAt || new Date().toISOString(),
          reconciled_with: t.reconciledWith || null,
          raw: t.raw || t,
        };
        const { error: uerr } = await supabase.from('transfers').upsert(up);
        if (uerr) console.error('upsert transfer error', uerr);
        // replace items
        const { error: derr } = await supabase.from('transfer_items').delete().eq('transfer_id', t.id);
        if (derr) console.error('delete transfer_items error', derr);
        const itemsToInsert = (t.items || []).map(it => ({ transfer_id: t.id, sku: it.sku, name: it.name, category: it.category, qty: it.qty, unit_price: it.unitPrice, line_total: it.lineTotal, metadata: it.metadata || null }));
        if (itemsToInsert.length) {
          const { error: ierr } = await supabase.from('transfer_items').insert(itemsToInsert);
          if (ierr) console.error('insert transfer_items error', ierr);
        }
      }
      return;
    }
    if (window.storage?.set) {
      await window.storage.set("transfers", JSON.stringify(transfers), true);
      return;
    }
    window.localStorage.setItem("store-transfer:transfers", JSON.stringify(transfers));
  } catch (e) {
    console.error("Failed to save transfers", e);
  }
}

async function loadProducts() {
  try {
    const SUPA_URL = import.meta.env.VITE_SUPABASE_URL;
    const SUPA_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
    const supabase = SUPA_URL && SUPA_KEY ? createClient(SUPA_URL, SUPA_KEY) : null;
    if (supabase) {
      const { data, error } = await supabase.from('products').select('*').order('sku', { ascending: true });
      if (error) throw error;
      if (!data) return null;
      return data.map(p => ({ sku: p.sku, name: p.name, category: p.category, subcategory: p.subcategory, unitPrice: p.unit_price ?? p.unitPrice, priceType: p.price_type, packQty: p.pack_qty ?? p.packQty, packPrice: p.pack_price ?? p.packPrice, metadata: p.metadata, createdAt: p.created_at }));
    }
    if (window.storage?.get) {
      const res = await window.storage.get("products", true);
      return res ? JSON.parse(res.value) : null;
    }
    return JSON.parse(window.localStorage.getItem("store-transfer:products") || "null");
  } catch (e) {
    console.error('loadProducts error', e);
    return null;
  }
}

async function saveProducts(products) {
  try {
    const SUPA_URL = import.meta.env.VITE_SUPABASE_URL;
    const SUPA_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
    const supabase = SUPA_URL && SUPA_KEY ? createClient(SUPA_URL, SUPA_KEY) : null;
    if (supabase) {
      const toUpsert = products.map(p => ({ sku: p.sku, name: p.name, category: p.category, subcategory: p.subcategory || null, unit_price: p.unitPrice, price_type: p.priceType, pack_qty: p.packQty || null, pack_price: p.packPrice || null, metadata: p.metadata || null, created_at: p.createdAt || new Date().toISOString() }));
      const { error } = await supabase.from('products').upsert(toUpsert);
      if (error) console.error('upsert products error', error);
      return;
    }
    if (window.storage?.set) {
      await window.storage.set("products", JSON.stringify(products), true);
      return;
    }
    window.localStorage.setItem("store-transfer:products", JSON.stringify(products));
  } catch (e) {
    console.error('saveProducts error', e);
  }
}

function exportToExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name);
  });
  XLSX.writeFile(wb, filename);
}

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');
`;

function Sidebar({ view, setView, stores, productCount }) {
  const items = [
    { id: "new", label: "New Transfer", icon: ArrowRightLeft },
    { id: "reconcile", label: "Reconcile", icon: CheckCircle2 },
    { id: "history", label: "Transfer History", icon: History },
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "products", label: "Product Database", icon: Package },
    { id: "addproducts", label: "Add Products", icon: ListPlus },
    { id: "invoiceexcel", label: "Invoice to Excel", icon: ScanLine },
  ];
  return (
    <div data-sidebar style={{ width: 220, background: "#151922", color: "#e8e6df", display: "flex", flexDirection: "column", flexShrink: 0 }}>
      <div style={{ padding: "22px 20px 18px 20px", borderBottom: "1px solid #262c3a" }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, letterSpacing: "-0.01em" }}>
          Store<span style={{ color: "#e8a33d" }}>Transfer</span>
        </div>
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#7b8194", marginTop: 3 }}>
          {stores.join("  ⇄  ")}
        </div>
      </div>
      <div style={{ padding: "14px 10px", flex: 1 }}>
        {items.map(it => {
          const Icon = it.icon;
          const active = view === it.id;
          return (
            <button
              key={it.id}
              onClick={() => setView(it.id)}
              style={{
                width: "100%", display: "flex", alignItems: "center", gap: 10,
                padding: "9px 12px", marginBottom: 3, borderRadius: 8, border: "none",
                background: active ? "#242a38" : "transparent",
                color: active ? "#f4a339" : "#c3c7d1",
                fontFamily: "'Inter', sans-serif", fontSize: 13.5, fontWeight: active ? 600 : 500,
                cursor: "pointer", textAlign: "left", transition: "background 0.15s"
              }}
            >
              <Icon size={16} strokeWidth={2} />
              {it.label}
            </button>
          );
        })}
      </div>
      <div style={{ padding: "14px 20px", borderTop: "1px solid #262c3a", fontFamily: "'IBM Plex Mono', monospace", fontSize: 10, color: "#5b6172" }}>
        {productCount} SKUs loaded
      </div>
    </div>
  );
}

function CategoryBadge({ category }) {
  const c = CATEGORY_COLORS[category] || { bg: "#f1f1f1", text: "#555", dot: "#999" };
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5, padding: "2.5px 8px",
      borderRadius: 20, background: c.bg, color: c.text, fontSize: 11, fontWeight: 600,
      fontFamily: "'Inter', sans-serif"
    }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: c.dot }} />
      {category}
    </span>
  );
}

function StatusPill({ status }) {
  const map = {
    matched: { bg: "#effaf3", text: "#166534", label: "Matched" },
    fuzzy: { bg: "#fff8e1", text: "#8a6116", label: "Fuzzy match — verify" },
    ambiguous: { bg: "#fff8e1", text: "#8a6116", label: "Duplicate SKU — verify" },
    unmatched: { bg: "#fdf2f2", text: "#a5202a", label: "Not found" },
  };
  const m = map[status] || map.unmatched;
  return (
    <span style={{
      padding: "2.5px 8px", borderRadius: 6, background: m.bg, color: m.text,
      fontSize: 11, fontWeight: 600, fontFamily: "'Inter', sans-serif", whiteSpace: "nowrap"
    }}>
      {m.label}
    </span>
  );
}

function ProductPicker({ value, onSelect, onClose, products }) {
  const [q, setQ] = useState("");
  const results = useMemo(() => {
    if (!q.trim()) return products.slice(0, 30);
    return products.filter(p => normalize(p.name).includes(normalize(q)) || normalize(p.sku).includes(normalize(q))).slice(0, 30);
  }, [q, products]);
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,17,23,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 12, width: 480, maxHeight: "70vh", display: "flex", flexDirection: "column", overflow: "hidden" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: 14, borderBottom: "1px solid #eee", display: "flex", alignItems: "center", gap: 8 }}>
          <Search size={15} color="#888" />
          <input autoFocus value={q} onChange={e => setQ(e.target.value)} placeholder="Search product name or SKU..."
            style={{ border: "none", outline: "none", flex: 1, fontFamily: "'Inter', sans-serif", fontSize: 13.5 }} />
          <X size={16} color="#888" style={{ cursor: "pointer" }} onClick={onClose} />
        </div>
        <div style={{ overflowY: "auto" }}>
          {results.map((p, i) => (
            <div key={p.sku + i} onClick={() => { onSelect(p); onClose(); }}
              style={{ padding: "9px 14px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #f5f5f5" }}
              onMouseEnter={e => e.currentTarget.style.background = "#fafafa"}
              onMouseLeave={e => e.currentTarget.style.background = "#fff"}
            >
              <div>
                <div style={{ fontSize: 13, fontWeight: 500, fontFamily: "'Inter', sans-serif" }}>{p.name}</div>
                <div style={{ fontSize: 11, color: "#999", fontFamily: "'IBM Plex Mono', monospace" }}>{p.sku} · {p.category}</div>
              </div>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, fontWeight: 600 }}>{money(p.unitPrice)}</div>
            </div>
          ))}
          {results.length === 0 && <div style={{ padding: 20, textAlign: "center", color: "#999", fontSize: 13 }}>No products found</div>}
        </div>
      </div>
    </div>
  );
}

function NewTransferView({ stores, products, onCreateTransfer, initialTransfer }) {
  const [entryMode, setEntryMode] = useState(initialTransfer ? "manual" : "upload");
  const [rows, setRows] = useState(() => initialTransfer ? initialTransfer.items.map((item, i) => ({
    id: uid(), rowNum: i + 1, sku: item.sku, inputName: item.name, qty: item.qty,
    product: products.find(p => p.sku === item.sku && p.name === item.name) || item, status: "matched",
  })) : []);
  const [fromStore, setFromStore] = useState(initialTransfer?.fromStore || stores[0]);
  const [toStore, setToStore] = useState(initialTransfer?.toStore || stores[1] || stores[0]);
  const [date, setDate] = useState(initialTransfer?.date || new Date().toISOString().slice(0, 10));
  const [notes, setNotes] = useState(initialTransfer?.notes || "");
  const [fileName, setFileName] = useState("");
  const [pickerIndex, setPickerIndex] = useState(null);
  const [error, setError] = useState("");
  const [invoice, setInvoice] = useState(null);
  const fileInputRef = useRef(null);
  const skuIndex = useMemo(() => buildSkuIndex(products), [products]);

  const handleFile = async (file) => {
    setError("");
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (json.length === 0) { setError("The sheet appears to be empty."); return; }
      const headers = Object.keys(json[0]);
      const skuKey = findHeaderKey(headers, ["SKU", "Product SKU"]);
      const nameKey = findHeaderKey(headers, ["Product Name", "Product", "Item", "Name"]);
      const qtyKey = findHeaderKey(headers, ["Quantity", "Qty"]);
      const fromKey = findHeaderKey(headers, ["From Store", "From"]);
      const toKey = findHeaderKey(headers, ["To Store", "To"]);
      if (!skuKey && !nameKey) {
        setError("Couldn't find a 'Product Name' or 'SKU' column in this sheet. Expected columns: Product Name, Quantity, From Store, To Store (SKU is optional).");
        return;
      }
      const parsedRows = json.map((r, i) => {
        const sku = skuKey ? String(r[skuKey]).trim() : "";
        const productName = nameKey ? String(r[nameKey]).trim() : "";
        const qty = qtyKey ? Number(r[qtyKey]) || 0 : 0;
        const rowFrom = fromKey && r[fromKey] ? String(r[fromKey]).trim() : "";
        const rowTo = toKey && r[toKey] ? String(r[toKey]).trim() : "";
        if (!sku && !productName) return null;
        const match = matchProduct({ sku, productName }, skuIndex, products);
        return {
          id: uid(), rowNum: i + 2, sku: sku || (match.product ? match.product.sku : ""),
          inputName: productName, qty: qty || 1,
          rowFrom, rowTo,
          product: match.product, status: match.status,
        };
      }).filter(Boolean);
      setRows(parsedRows);
      if (fromKey || toKey) {
        const firstFrom = parsedRows.find(r => r.rowFrom)?.rowFrom;
        const firstTo = parsedRows.find(r => r.rowTo)?.rowTo;
        if (firstFrom && stores.includes(firstFrom)) setFromStore(firstFrom);
        if (firstTo && stores.includes(firstTo)) setToStore(firstTo);
      }
    } catch (e) {
      setError("Couldn't read this file. Make sure it's a valid .xlsx or .csv file.");
    }
  };

  const updateRow = (id, patch) => setRows(rs => rs.map(r => r.id === id ? { ...r, ...patch } : r));
  const removeRow = (id) => setRows(rs => rs.filter(r => r.id !== id));
  const addManualItem = () => {
    setEntryMode("manual");
    setRows(rs => [...rs, { id: uid(), rowNum: rs.length + 1, sku: "", inputName: "", qty: 1, product: null, status: "unmatched" }]);
  };

  const matchedCount = rows.filter(r => r.product).length;
  const totalRows = rows.length;
  const subtotal = rows.reduce((sum, r) => r.product ? sum + r.product.unitPrice * r.qty : sum, 0);
  const hst = subtotal * HST_RATE;
  const total = subtotal + hst;

  const canGenerate = rows.length > 0
    && rows.every(r => r.product && r.status === "matched" && Number.isInteger(r.qty) && r.qty > 0)
    && fromStore !== toStore;

  const handleGenerate = () => {
    const items = rows.map(r => ({
      sku: r.product.sku, name: r.product.name, category: r.product.category,
      qty: r.qty, unitPrice: r.product.unitPrice, lineTotal: round2(r.product.unitPrice * r.qty),
    }));
    const t = {
      id: initialTransfer?.id || uid(), date, fromStore, toStore, notes, items,
      subtotal: round2(subtotal), hst: round2(hst), total: round2(total),
      createdAt: initialTransfer?.createdAt || new Date().toISOString(),
    };
    onCreateTransfer(t);
    setInvoice(t);
  };

  const reset = () => {
    setRows([]); setFileName(""); setInvoice(null); setNotes(""); setEntryMode("upload");
    setDate(new Date().toISOString().slice(0, 10));
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  function round2(n) { return Math.round(n * 100) / 100; }

  if (invoice) {
    return <InvoiceView invoice={invoice} onNewTransfer={reset} />;
  }

  return (
    <div style={{ padding: "28px 36px", maxWidth: 980 }}>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>New Transfer</h1>
      <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 24 }}>{initialTransfer ? "Update the transfer, then regenerate the invoice." : "Upload a transfer sheet or add items manually, review the matches, and generate the invoice."}</p>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 22 }}>
        <Field label="From Store">
          <select value={fromStore} onChange={e => setFromStore(e.target.value)} style={selectStyle}>
            {stores.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="To Store">
          <select value={toStore} onChange={e => setToStore(e.target.value)} style={selectStyle}>
            {stores.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Transfer Date">
          <input type="date" value={date} onChange={e => setDate(e.target.value)} style={selectStyle} />
        </Field>
      </div>
      {fromStore === toStore && <div style={{ color: "#a5202a", fontSize: 12.5, marginTop: -12, marginBottom: 16 }}>From and To store must be different.</div>}

      {!initialTransfer && rows.length === 0 && (
        <div style={{ display: "flex", gap: 6, marginBottom: 18 }}>
          {[['upload', 'Upload Excel', Upload], ['manual', 'Add Manually', ListPlus]].map(([id, label, Icon]) => (
            <button key={id} onClick={() => { setEntryMode(id); if (id === "manual" && rows.length === 0) addManualItem(); }}
              style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 8, border: entryMode === id ? "1px solid #171a21" : "1px solid #e8e9ee", background: entryMode === id ? "#171a21" : "#fff", color: entryMode === id ? "#fff" : "#3a3f4c", fontSize: 12.5, fontWeight: 600, cursor: "pointer", fontFamily: "'Inter', sans-serif" }}>
              <Icon size={13} /> {label}
            </button>
          ))}
        </div>
      )}

      {rows.length === 0 && entryMode === "upload" ? (
        <div
          onDragOver={e => e.preventDefault()}
          onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: "2px dashed #d8dbe3", borderRadius: 12, padding: "44px 20px", textAlign: "center",
            cursor: "pointer", background: "#fbfbfc"
          }}
        >
          <Upload size={26} color="#a3a8b8" style={{ marginBottom: 10 }} />
          <div style={{ fontFamily: "'Inter', sans-serif", fontWeight: 600, fontSize: 14.5, color: "#3a3f4c" }}>Drop a transfer sheet here, or click to browse</div>
          <div style={{ fontSize: 12, color: "#9aa0ae", marginTop: 5 }}>.xlsx or .csv — columns: Product Name, Quantity, From Store, To Store</div>
          <div style={{ fontSize: 11.5, color: "#bdc1cc", marginTop: 3 }}>SKU is optional — matching goes by product name, so it's fine if the SKU is missing or doesn't line up exactly.</div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={e => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
      ) : (
        <>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "#666" }}>
              {entryMode === "upload" ? <FileSpreadsheet size={14} /> : <ListPlus size={14} />} {fileName || "Manual transfer"} — {totalRows} rows, {matchedCount} matched
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={addManualItem} style={ghostBtn}>Add item</button>
              <button onClick={reset} style={ghostBtn}>Start over</button>
            </div>
          </div>

          <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", marginBottom: 18 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: "#f8f8fa", textAlign: "left" }}>
                  {["Row", "Matched Product", "SKU", "Category", "Qty", "Unit Price", "Line Total", "Status", ""].map(h => (
                    <th key={h} style={{ padding: "8px 10px", fontWeight: 600, color: "#6b7080", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.03em" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map(r => (
                  <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4" }}>
                    <td style={{ padding: "8px 10px", color: "#999" }}>{r.rowNum}</td>
                    <td style={{ padding: "8px 10px", fontWeight: 500 }}>
                      {r.product ? r.product.name : (r.inputName || r.sku || "—")}
                      {r.status !== "matched" && (
                        <button onClick={() => setPickerIndex(r.id)} style={{ ...ghostBtn, marginLeft: 8, padding: "1px 7px" }}>Fix match</button>
                      )}
                    </td>
                    <td style={{ padding: "8px 10px", fontFamily: "'IBM Plex Mono', monospace", color: "#666" }}>{r.product ? r.product.sku : r.sku || "—"}</td>
                    <td style={{ padding: "8px 10px" }}>{r.product ? <CategoryBadge category={r.product.category} /> : "—"}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <input type="number" min="1" step="1" value={r.qty} onChange={e => updateRow(r.id, { qty: Number(e.target.value) || 0 })}
                        style={{ width: 52, padding: "3px 6px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace" }} />
                    </td>
                    <td style={{ padding: "8px 10px", fontFamily: "'IBM Plex Mono', monospace" }}>{r.product ? money(r.product.unitPrice) : "—"}</td>
                    <td style={{ padding: "8px 10px", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{r.product ? money(r.product.unitPrice * r.qty) : "—"}</td>
                    <td style={{ padding: "8px 10px" }}><StatusPill status={r.status} /></td>
                    <td style={{ padding: "8px 10px" }}>
                      <Trash2 size={14} color="#c3c7d1" style={{ cursor: "pointer" }} onClick={() => removeRow(r.id)} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {rows.some(r => !r.product || r.status !== "matched" || !Number.isInteger(r.qty) || r.qty <= 0) && (
            <div style={{ display: "flex", gap: 8, alignItems: "flex-start", background: "#fff8e1", border: "1px solid #f5e2a8", borderRadius: 8, padding: "10px 14px", marginBottom: 18, fontSize: 12.5, color: "#8a6116" }}>
              <AlertTriangle size={15} style={{ flexShrink: 0, marginTop: 1 }} />
              Review every row before generating: choose a product for unmatched, fuzzy, or duplicate matches, and enter a whole quantity of at least 1.
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <div style={{ width: 260, fontSize: 13.5 }}>
              <SummaryLine label="Subtotal" value={money(subtotal)} />
              <SummaryLine label="HST (13%)" value={money(hst)} />
              <SummaryLine label="Total" value={money(total)} bold />
            </div>
          </div>

          <div style={{ marginTop: 16 }}>
            <Field label="Notes (optional)">
              <input value={notes} onChange={e => setNotes(e.target.value)} placeholder="e.g. weekly restock transfer"
                style={{ ...selectStyle, width: "100%" }} />
            </Field>
          </div>

          {error && <div style={{ color: "#a5202a", fontSize: 12.5, marginTop: 10 }}>{error}</div>}

          <button
            disabled={!canGenerate}
            onClick={handleGenerate}
            style={{
              marginTop: 20, background: canGenerate ? "#171a21" : "#d8dbe3", color: "#fff", border: "none",
              padding: "11px 22px", borderRadius: 9, fontSize: 14, fontWeight: 600, cursor: canGenerate ? "pointer" : "not-allowed",
              fontFamily: "'Inter', sans-serif", display: "flex", alignItems: "center", gap: 6
            }}
          >
            Generate Invoice <ChevronRight size={15} />
          </button>
        </>
      )}

      {error && rows.length === 0 && <div style={{ color: "#a5202a", fontSize: 12.5, marginTop: 12 }}>{error}</div>}

      {pickerIndex && (
        <ProductPicker
          products={products}
          onClose={() => setPickerIndex(null)}
          onSelect={(p) => updateRow(pickerIndex, { product: p, status: "matched", sku: p.sku })}
        />
      )}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div>
      <div style={{ fontSize: 11.5, fontWeight: 600, color: "#767c8c", marginBottom: 5, textTransform: "uppercase", letterSpacing: "0.03em" }}>{label}</div>
      {children}
    </div>
  );
}

function SummaryLine({ label, value, bold }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", borderTop: bold ? "1px solid #e8e9ee" : "none", marginTop: bold ? 5 : 0 }}>
      <span style={{ color: bold ? "#171a21" : "#767c8c", fontWeight: bold ? 700 : 500 }}>{label}</span>
      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: bold ? 700 : 500 }}>{value}</span>
    </div>
  );
}

const selectStyle = {
  padding: "8px 10px", border: "1px solid #d8dbe3", borderRadius: 8, fontSize: 13,
  fontFamily: "'Inter', sans-serif", width: "100%", boxSizing: "border-box", background: "#fff"
};
const ghostBtn = {
  background: "none", border: "1px solid #d8dbe3", borderRadius: 6, padding: "3px 9px",
  fontSize: 11.5, cursor: "pointer", color: "#3a3f4c", fontFamily: "'Inter', sans-serif"
};

function InvoiceView({ invoice, onNewTransfer }) {
  return (
    <div style={{ padding: "28px 36px" }}>
      <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20, maxWidth: 640 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#166534", fontWeight: 600, fontSize: 14 }}>
          <CheckCircle2 size={17} /> Invoice generated and saved to Transfer History
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => window.print()} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 5, padding: "6px 12px" }}><Printer size={13} /> Print</button>
          <button onClick={onNewTransfer} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 5, padding: "6px 12px" }}><PlusCircle size={13} /> New Transfer</button>
        </div>
      </div>
      <InvoiceSlip invoice={invoice} />
    </div>
  );
}

function InvoiceSlip({ invoice }) {
  return (
    <div style={{
      maxWidth: 640, background: "#fff", border: "1px solid #e8e9ee", borderRadius: 12,
      padding: "30px 34px", fontFamily: "'Inter', sans-serif", position: "relative"
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: "2px dashed #e0e2e8", paddingBottom: 18, marginBottom: 18 }}>
        <div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 19 }}>Transfer Invoice</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: "#999", marginTop: 3 }}>#{invoice.id.toUpperCase()}</div>
        </div>
        <div style={{ textAlign: "right", fontSize: 12.5, color: "#666" }}>{invoice.date}</div>
      </div>

      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 22 }}>
        <StoreChip label="FROM" name={invoice.fromStore} color="#a5202a" />
        <ArrowRightLeft size={16} color="#c3c7d1" />
        <StoreChip label="TO" name={invoice.toStore} color="#166534" />
      </div>

      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5, marginBottom: 18 }}>
        <thead>
          <tr style={{ borderBottom: "1px solid #e8e9ee", textAlign: "left" }}>
            <th style={{ padding: "0 0 8px 0", fontSize: 10.5, color: "#999", textTransform: "uppercase", letterSpacing: "0.03em" }}>Item</th>
            <th style={{ padding: "0 0 8px 0", fontSize: 10.5, color: "#999", textTransform: "uppercase", letterSpacing: "0.03em", textAlign: "center" }}>Qty</th>
            <th style={{ padding: "0 0 8px 0", fontSize: 10.5, color: "#999", textTransform: "uppercase", letterSpacing: "0.03em", textAlign: "right" }}>Unit</th>
            <th style={{ padding: "0 0 8px 0", fontSize: 10.5, color: "#999", textTransform: "uppercase", letterSpacing: "0.03em", textAlign: "right" }}>Total</th>
          </tr>
        </thead>
        <tbody>
          {invoice.items.map((it, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #f5f5f7" }}>
              <td style={{ padding: "8px 0" }}>
                <div style={{ fontWeight: 500 }}>{it.name}</div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 10.5, color: "#aaa" }}>{it.sku}</div>
              </td>
              <td style={{ padding: "8px 0", textAlign: "center", fontFamily: "'IBM Plex Mono', monospace" }}>{it.qty}</td>
              <td style={{ padding: "8px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", color: "#666" }}>{money(it.unitPrice)}</td>
              <td style={{ padding: "8px 0", textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{money(it.lineTotal)}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <div style={{ width: 210 }}>
          <SummaryLine label="Subtotal" value={money(invoice.subtotal)} />
          <SummaryLine label="HST (13%)" value={money(invoice.hst)} />
          <SummaryLine label="Total" value={money(invoice.total)} bold />
        </div>
      </div>
      {invoice.notes && (
        <div style={{ marginTop: 18, paddingTop: 14, borderTop: "2px dashed #e0e2e8", fontSize: 12, color: "#767c8c" }}>
          <strong style={{ color: "#3a3f4c" }}>Notes:</strong> {invoice.notes}
        </div>
      )}
    </div>
  );
}

function StoreChip({ label, name, color }) {
  return (
    <div style={{ textAlign: "center" }}>
      <div style={{ fontSize: 9.5, fontWeight: 700, color: "#aaa", letterSpacing: "0.06em" }}>{label}</div>
      <div style={{
        marginTop: 3, padding: "5px 14px", borderRadius: 20, background: color + "14", color,
        fontWeight: 700, fontSize: 13.5, fontFamily: "'Space Grotesk', sans-serif"
      }}>{name}</div>
    </div>
  );
}

function HistoryView({ transfers, loading, onEdit, onDelete }) {
  const [selected, setSelected] = useState(null);
  if (selected) {
    return (
      <div style={{ padding: "28px 36px" }}>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          <button onClick={() => setSelected(null)} style={ghostBtn}>← Back to history</button>
          <button onClick={() => onEdit(selected)} style={ghostBtn}>Edit transfer</button>
          <button onClick={() => { if (window.confirm("Delete this transfer? This cannot be undone.")) { onDelete(selected.id); setSelected(null); } }} style={{ ...ghostBtn, color: "#a5202a" }}>Delete</button>
        </div>
        <InvoiceSlip invoice={selected} />
      </div>
    );
  }
  const handleExport = () => {
    const summaryRows = transfers.map(t => ({
      "Transfer ID": t.id, "Date": t.date, "From Store": t.fromStore, "To Store": t.toStore,
      "Item Count": t.items.length, "Subtotal": t.subtotal, "HST": t.hst, "Total": t.total, "Notes": t.notes || ""
    }));
    const lineItemRows = transfers.flatMap(t => t.items.map(it => ({
      "Transfer ID": t.id, "Date": t.date, "From Store": t.fromStore, "To Store": t.toStore,
      "SKU": it.sku, "Product Name": it.name, "Category": it.category, "Qty": it.qty,
      "Unit Price": it.unitPrice, "Line Total": it.lineTotal
    })));
    exportToExcel("transfer_history.xlsx", [
      { name: "Transfers", rows: summaryRows },
      { name: "Line Items", rows: lineItemRows },
    ]);
  };

  return (
    <div style={{ padding: "28px 36px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Transfer History</h1>
          <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 22 }}>{transfers.length} transfer{transfers.length !== 1 ? "s" : ""} recorded</p>
        </div>
        {transfers.length > 0 && (
          <button onClick={handleExport} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 6, padding: "7px 13px" }}>
            <Download size={13} /> Export as Excel
          </button>
        )}
      </div>
      {loading ? (
        <div style={{ color: "#999", fontSize: 13 }}>Loading...</div>
      ) : transfers.length === 0 ? (
        <EmptyState text="No transfers yet. Generate one from the New Transfer tab." />
      ) : (
        <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden" }}>
          {transfers.slice().reverse().map(t => (
            <div key={t.id} onClick={() => setSelected(t)}
              style={{ padding: "13px 16px", borderBottom: "1px solid #f0f1f4", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}
              onMouseEnter={e => e.currentTarget.style.background = "#fafafa"}
              onMouseLeave={e => e.currentTarget.style.background = "#fff"}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: "#aaa", width: 82 }}>{t.date}</div>
                <div style={{ fontSize: 13, fontWeight: 600, color: "#171a21", display: "flex", alignItems: "center", gap: 6 }}>
                  {t.fromStore} <ArrowRightLeft size={12} color="#c3c7d1" /> {t.toStore}
                </div>
                <div style={{ fontSize: 12, color: "#999" }}>{t.items.length} item{t.items.length !== 1 ? "s" : ""}</div>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5 }}>{money(t.total)}</div>
                <button onClick={e => { e.stopPropagation(); onEdit(t); }} style={{ ...ghostBtn, padding: "4px 8px" }}>Edit</button>
                <button onClick={e => { e.stopPropagation(); if (window.confirm("Delete this transfer? This cannot be undone.")) onDelete(t.id); }} style={{ ...ghostBtn, padding: "4px 8px", color: "#a5202a" }}>Delete</button>
                <ChevronRight size={15} color="#c3c7d1" />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ReconcileView({ transfers, products, stores, onCreateTransfers }) {
  const prodBySku = useMemo(() => {
    const m = new Map(); products.forEach(p => m.set(p.sku, p)); return m;
  }, [products]);

  const computed = useMemo(() => {
    const pairs = {};
    for (const t of transfers) {
      for (const it of t.items) {
        const sku = it.sku || it.name || "";
        const a = t.fromStore; const b = t.toStore;
        const [s1, s2] = a < b ? [a, b] : [b, a];
        const key = `${sku}||${s1}||${s2}`;
        const sign = (t.fromStore === s1 && t.toStore === s2) ? 1 : -1;
        if (!pairs[key]) pairs[key] = { sku, name: it.name || (prodBySku.get(sku) && prodBySku.get(sku).name) || "", unitPrice: it.unitPrice || (prodBySku.get(sku) && prodBySku.get(sku).unitPrice) || 0, s1, s2, delta: 0, sources: [] };
        const qty = Number(it.qty) || 0;
        pairs[key].delta += sign * qty;
        pairs[key].sources.push({ transferId: t.id, date: t.date, fromStore: t.fromStore, toStore: t.toStore, qty: sign * qty, unitPrice: it.unitPrice || pairs[key].unitPrice });
      }
    }
    const rows = [];
    Object.values(pairs).forEach(p => {
      if (!p.delta) return;
      if (p.delta > 0) rows.push({ id: uid(), sku: p.sku, name: p.name, from: p.s1, to: p.s2, qty: p.delta, unitPrice: p.unitPrice, sources: p.sources });
      else rows.push({ id: uid(), sku: p.sku, name: p.name, from: p.s2, to: p.s1, qty: -p.delta, unitPrice: p.unitPrice, sources: p.sources });
    });
    return rows;
  }, [transfers, prodBySku]);

  const [rows, setRows] = useState(computed);
  useEffect(() => setRows(computed), [computed]);
  const [detail, setDetail] = useState(null);
  const [generated, setGenerated] = useState([]);
  const [genIndex, setGenIndex] = useState(0);

  const updateQty = (id, qty) => setRows(rs => rs.map(r => r.id === id ? { ...r, qty: Math.max(0, Math.floor(Number(qty) || 0)) } : r));

  const groupedByPair = useMemo(() => {
    const m = new Map();
    for (const r of rows.filter(x => x.qty > 0)) {
      const key = `${r.from}||${r.to}`;
      if (!m.has(key)) m.set(key, []);
      m.get(key).push(r);
    }
    return m;
  }, [rows]);

  const [generatedMapping, setGeneratedMapping] = useState({});

  const handleCreate = () => {
    if (groupedByPair.size === 0) return;
    const today = new Date().toISOString().slice(0,10);
    const created = [];
    const markMap = {};
    for (const [key, items] of groupedByPair.entries()) {
      const [fromStore, toStore] = key.split("||");
      const tItems = items.map(it => ({ sku: it.sku, name: it.name, qty: it.qty, unitPrice: it.unitPrice || 0, lineTotal: Math.round((it.unitPrice || 0) * it.qty * 100)/100 }));
      const subtotal = tItems.reduce((s,it) => s + (Number(it.lineTotal)||0), 0);
      const hst = Math.round(subtotal * HST_RATE * 100)/100;
      const total = Math.round((subtotal + hst) * 100)/100;
      const t = { id: uid(), date: today, fromStore, toStore, notes: `Reconciled invoice (${today})`, items: tItems, subtotal, hst, total, createdAt: new Date().toISOString() };
      created.push(t);
      // collect source transfer ids for these items
      const invoiceId = t.id;
      const sourceIds = new Set();
      items.forEach(it => (it.sources || []).forEach(s => sourceIds.add(s.transferId)));
      sourceIds.forEach(sid => { markMap[sid] = markMap[sid] || []; markMap[sid].push(invoiceId); });
    }
    if (created.length) {
      setGenerated(created);
      setGeneratedMapping(markMap);
      setGenIndex(0);
    }
  };

  return (
    <div style={{ padding: "28px 36px" }}>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Reconcile Transfers</h1>
      <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 22 }}>See net movements between store pairs. Edit quantities and create adjustment transfers to reconcile back-and-forth moves.</p>
      {rows.length === 0 ? (
        <EmptyState text="No net movements to reconcile." />
      ) : (
        <div>
          <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", maxHeight: 520, overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead style={{ position: "sticky", top: 0, background: "#f8f8fa", zIndex: 1 }}>
                <tr style={{ textAlign: "left" }}>
                  {["SKU", "Product", "From", "To", "Qty"].map(h => (
                    <th key={h} style={{ padding: "8px 12px", fontWeight: 600, color: "#6b7080", fontSize: 11, textTransform: "uppercase" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map(r => (
                  <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4", cursor: "pointer" }} onClick={() => setDetail(r)}>
                    <td style={{ padding: "10px" }}>{r.sku || "—"}</td>
                    <td style={{ padding: "10px", fontWeight: 600 }}>{r.name || "—"}</td>
                    <td style={{ padding: "10px" }}>{r.from}</td>
                    <td style={{ padding: "10px" }}>{r.to}</td>
                    <td style={{ padding: "10px" }}>
                      <input type="number" value={r.qty} min={0} onChange={e => updateQty(r.id, e.target.value)} style={{ width: 96, padding: 6 }} onClick={e => e.stopPropagation()} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: 14, display: "flex", gap: 10, alignItems: "center" }}>
            <button onClick={handleCreate} style={{ background: "#171a21", color: "#fff", border: "none", padding: "10px 16px", borderRadius: 8, cursor: "pointer" }}>Generate Invoice(s)</button>
            <button onClick={() => setRows(computed)} style={{ ...ghostBtn }}>Reset</button>
            <div style={{ color: "#767c8c", fontSize: 13 }}>Click a product row to see contributing transactions.</div>
          </div>
        </div>
      )}
      {detail && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,17,23,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 120 }} onClick={() => setDetail(null)}>
          <div style={{ background: "#fff", borderRadius: 12, width: 720, maxHeight: "70vh", overflowY: "auto", padding: 18 }} onClick={e => e.stopPropagation()}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontWeight: 700 }}>{detail.name || detail.sku}</div>
              <button onClick={() => setDetail(null)} style={ghostBtn}>Close</button>
            </div>
            <div style={{ marginBottom: 10, color: "#666" }}>Transactions contributing to this net movement (signed quantities show direction):</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ textAlign: "left", borderBottom: "1px solid #eee" }}>
                  { ["Date","Transfer ID","From","To","Qty","Unit"] .map(h => <th key={h} style={{ padding: 8, color: "#666", fontSize: 12 }}>{h}</th>) }
                </tr>
              </thead>
              <tbody>
                {(detail.sources||[]).map((s,i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #f5f5f7" }}>
                    <td style={{ padding: 8, fontFamily: "'IBM Plex Mono', monospace" }}>{s.date}</td>
                    <td style={{ padding: 8, fontFamily: "'IBM Plex Mono', monospace" }}>{s.transferId}</td>
                    <td style={{ padding: 8 }}>{s.fromStore}</td>
                    <td style={{ padding: 8 }}>{s.toStore}</td>
                    <td style={{ padding: 8, fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{s.qty}</td>
                    <td style={{ padding: 8, fontFamily: "'IBM Plex Mono', monospace" }}>{money(s.unitPrice)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {generated.length > 0 && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,17,23,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 130 }}>
          <div style={{ width: 760, maxHeight: "90vh", overflowY: "auto", background: "#fff", borderRadius: 12, padding: 18 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontWeight: 700 }}>Generated Invoices ({generated.length})</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => { setGenerated([]); }} style={ghostBtn}>Close</button>
              </div>
            </div>
            <div>
              <div style={{ marginBottom: 10, display: "flex", gap: 8, alignItems: "center" }}>
                <button onClick={() => setGenIndex(i => Math.max(0, i-1))} disabled={genIndex===0} style={ghostBtn}>Prev</button>
                <button onClick={() => setGenIndex(i => Math.min(generated.length-1, i+1))} disabled={genIndex===generated.length-1} style={ghostBtn}>Next</button>
                <div style={{ marginLeft: 8, color: "#666" }}>Viewing {genIndex+1} of {generated.length}</div>
                <button onClick={() => window.print()} style={{ marginLeft: 'auto', ...ghostBtn }}>Print</button>
                <button onClick={() => {
                  // save generated invoices and mark source transfers
                  if (typeof onCreateTransfers === 'function') onCreateTransfers(generated);
                  if (typeof onMarkTransfers === 'function' && generatedMapping && Object.keys(generatedMapping).length) onMarkTransfers(generatedMapping);
                  setGenerated([]);
                  setGeneratedMapping({});
                  alert('Saved generated invoices to Transfer History and marked source transfers as reconciled.');
                }} style={{ ...ghostBtn, marginLeft: 8 }}>Save to History</button>
              </div>
              <InvoiceSlip invoice={generated[genIndex]} />
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ border: "1px dashed #d8dbe3", borderRadius: 10, padding: "40px 20px", textAlign: "center", color: "#9aa0ae", fontSize: 13 }}>
      {text}
    </div>
  );
}

function DashboardView({ transfers, stores }) {
  const stats = useMemo(() => {
    const s = {};
    stores.forEach(store => { s[store] = { sentUnits: 0, receivedUnits: 0, sentValue: 0, receivedValue: 0 }; });
    transfers.forEach(t => {
      const units = t.items.reduce((sum, it) => sum + it.qty, 0);
      if (s[t.fromStore]) { s[t.fromStore].sentUnits += units; s[t.fromStore].sentValue += t.total; }
      if (s[t.toStore]) { s[t.toStore].receivedUnits += units; s[t.toStore].receivedValue += t.total; }
    });
    return s;
  }, [transfers, stores]);

  return (
    <div style={{ padding: "28px 36px" }}>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Dashboard</h1>
      <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 24 }}>Live totals across all recorded transfers.</p>
      <div style={{ display: "grid", gridTemplateColumns: `repeat(${stores.length}, 1fr)`, gap: 18 }}>
        {stores.map(store => {
          const d = stats[store];
          const net = d.receivedUnits - d.sentUnits;
          return (
            <div key={store} style={{ border: "1px solid #e8e9ee", borderRadius: 12, padding: "20px 22px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
                <Building2 size={16} color="#e8a33d" />
                <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16 }}>{store}</div>
              </div>
              <StatRow label="Units sent out" value={d.sentUnits} />
              <StatRow label="Units received" value={d.receivedUnits} />
              <StatRow label="Net units" value={(net >= 0 ? "+" : "") + net} highlight={net !== 0} />
              <div style={{ height: 1, background: "#f0f1f4", margin: "10px 0" }} />
              <StatRow label="Value sent out" value={money(d.sentValue)} />
              <StatRow label="Value received" value={money(d.receivedValue)} />
            </div>
          );
        })}
      </div>
    </div>
  );
}

function StatRow({ label, value, highlight }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "5px 0", fontSize: 13 }}>
      <span style={{ color: "#767c8c" }}>{label}</span>
      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600, color: highlight ? "#e8a33d" : "#171a21" }}>{value}</span>
    </div>
  );
}

function ProductsView({ products, onUpdateProduct }) {
  const [q, setQ] = useState("");
  const [cat, setCat] = useState("All");
  const categories = ["All", ...Array.from(new Set(products.map(p => p.category)))];
  const filtered = useMemo(() => {
    return products.filter(p => {
      if (cat !== "All" && p.category !== cat) return false;
      if (q && !normalize(p.name).includes(normalize(q)) && !normalize(p.sku).includes(normalize(q))) return false;
      return true;
    });
  }, [q, cat, products]);

  const handleExport = () => {
    const rows = products.map(p => ({
      "SKU": p.sku, "Product Name": p.name, "Category": p.category,
      "Unit Price": p.unitPrice, "Price Type": p.priceType, "Pack Qty": p.packQty || 1
    }));
    exportToExcel("product_database.xlsx", [{ name: "Products", rows }]);
  };
  const editProduct = (product, patch) => onUpdateProduct(products.indexOf(product), patch);

  return (
    <div style={{ padding: "28px 36px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Product Database</h1>
          <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 20 }}>{products.length} products across {categories.length - 1} categories — this is the source of truth for invoice pricing.</p>
        </div>
        <button onClick={handleExport} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 6, padding: "7px 13px" }}>
          <Download size={13} /> Export as Excel
        </button>
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
        <div style={{ position: "relative", flex: 1, maxWidth: 320 }}>
          <Search size={14} color="#aaa" style={{ position: "absolute", left: 10, top: 10 }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search by name or SKU..."
            style={{ ...selectStyle, paddingLeft: 30 }} />
        </div>
        <select value={cat} onChange={e => setCat(e.target.value)} style={{ ...selectStyle, width: 180 }}>
          {categories.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>
      <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", maxHeight: 560, overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
          <thead style={{ position: "sticky", top: 0, background: "#f8f8fa", zIndex: 1 }}>
            <tr style={{ textAlign: "left" }}>
              {["SKU", "Product Name", "Category", "Unit Price", "Price Type", "Pack Qty"].map(h => (
                <th key={h} style={{ padding: "8px 12px", fontWeight: 600, color: "#6b7080", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.03em" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.slice(0, 300).map((p, i) => (
              <tr key={p.sku + i} style={{ borderTop: "1px solid #f0f1f4" }}>
                <td style={{ padding: "6px 8px" }}><input value={p.sku} onChange={e => editProduct(p, { sku: e.target.value })} aria-label={`SKU for ${p.name}`} style={{ width: 104, padding: "4px 6px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }} /></td>
                <td style={{ padding: "6px 8px" }}><input value={p.name} onChange={e => editProduct(p, { name: e.target.value })} aria-label={`Name for ${p.sku}`} style={{ width: 210, padding: "4px 6px", border: "1px solid #ddd", borderRadius: 5, fontSize: 12 }} /></td>
                <td style={{ padding: "6px 8px" }}><select value={p.category} onChange={e => editProduct(p, { category: e.target.value })} aria-label={`Category for ${p.name}`} style={{ padding: "4px 5px", border: "1px solid #ddd", borderRadius: 5, fontSize: 11.5 }}>{CATEGORY_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}</select></td>
                <td style={{ padding: "6px 8px" }}><input type="number" min="0" step="0.01" value={p.unitPrice} onChange={e => editProduct(p, { unitPrice: Number(e.target.value) || 0 })} aria-label={`Price for ${p.name}`} style={{ width: 80, padding: "4px 6px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }} /></td>
                <td style={{ padding: "6px 8px" }}><select value={p.priceType || "Per Unit"} onChange={e => editProduct(p, { priceType: e.target.value })} aria-label={`Price type for ${p.name}`} style={{ padding: "4px 5px", border: "1px solid #ddd", borderRadius: 5, fontSize: 11.5 }}><option value="Per Unit">Per Unit</option><option value="Pack Price">Pack Price</option></select></td>
                <td style={{ padding: "6px 8px" }}><input type="number" min="1" step="1" value={p.packQty || 1} onChange={e => editProduct(p, { packQty: Number(e.target.value) || 1 })} aria-label={`Pack quantity for ${p.name}`} style={{ width: 58, padding: "4px 6px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5 }} /></td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length > 300 && (
          <div style={{ padding: 10, textAlign: "center", fontSize: 11.5, color: "#aaa" }}>Showing first 300 of {filtered.length} results — refine your search.</div>
        )}
      </div>
    </div>
  );
}

const CATEGORY_OPTIONS = ["Vape", "Convenience", "Cigarettes", "Beer & Alcohol", "Other"];

function blankProductRow() {
  return { id: uid(), sku: "", name: "", category: "Convenience", unitPrice: "", priceType: "Per Unit", packQty: 1 };
}

function parseInvoiceText(text, existingSkus) {
  return text.split(/\r?\n/).map(line => line.replace(/\s+/g, " ").trim()).filter(Boolean).map(line => {
    const priceMatches = [...line.matchAll(/\$?\d{1,6}(?:,\d{3})*(?:\.\d{2})/g)].map(m => Number(m[0].replace(/[$,]/g, "")));
    const skuMatch = line.match(/\b[A-Z0-9][A-Z0-9/_-]{2,}\b/);
    const qtyMatch = line.match(/(?:qty|quantity|x)\s*[:#]?\s*(\d+)/i);
    const price = priceMatches.length ? priceMatches[priceMatches.length - 1] : NaN;
    const sku = skuMatch ? skuMatch[0] : "";
    const name = line.replace(sku, "").replace(/(?:qty|quantity|x)\s*[:#]?\s*\d+/i, "").replace(/\$?\d{1,6}(?:,\d{3})*(?:\.\d{2})/g, "").replace(/[|;,]+/g, " ").trim();
    const valid = !!name && !isNaN(price) && !/^(subtotal|total|tax|hst|invoice|date|supplier|payment)/i.test(name);
    return { id: uid(), sku, name, category: "Other", unitPrice: price, priceType: "Per Unit", packQty: 1, qty: qtyMatch ? Number(qtyMatch[1]) : 1, valid, isDuplicate: existingSkus.has(sku), include: valid && !existingSkus.has(sku), needsReview: true };
  }).filter(r => r.name || r.sku);
}

function AddProductsView({ products, onAddProducts }) {
  const [mode, setMode] = useState("manual");
  const [manualRows, setManualRows] = useState([blankProductRow(), blankProductRow(), blankProductRow()]);
  const [savedCount, setSavedCount] = useState(null);
  const [bulkRows, setBulkRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [invoiceRows, setInvoiceRows] = useState([]);
  const [invoiceProgress, setInvoiceProgress] = useState("");
  const fileInputRef = useRef(null);
  const existingSkus = useMemo(() => new Set(products.map(p => p.sku)), [products]);

  const updateManualRow = (id, patch) => setManualRows(rs => rs.map(r => r.id === id ? { ...r, ...patch } : r));
  const removeManualRow = (id) => setManualRows(rs => rs.filter(r => r.id !== id));
  const addManualRow = () => setManualRows(rs => [...rs, blankProductRow()]);

  const manualValidRows = manualRows.filter(r => r.sku.trim() && r.name.trim() && r.unitPrice !== "" && !isNaN(Number(r.unitPrice)));

  const handleSaveManual = () => {
    const toAdd = manualValidRows.map(r => ({
      sku: r.sku.trim(), name: r.name.trim(), category: r.category,
      unitPrice: Math.round(Number(r.unitPrice) * 100) / 100, priceType: r.priceType, packQty: Number(r.packQty) || 1
    }));
    if (toAdd.length === 0) return;
    onAddProducts(toAdd);
    setSavedCount(toAdd.length);
    setManualRows([blankProductRow(), blankProductRow(), blankProductRow()]);
  };

  const handleBulkFile = async (file) => {
    setError(""); setFileName(file.name); setSavedCount(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (json.length === 0) { setError("The sheet appears to be empty."); return; }
      const headers = Object.keys(json[0]);
      const skuKey = findHeaderKey(headers, ["SKU"]);
      const nameKey = findHeaderKey(headers, ["Product Name", "Name"]);
      const priceKey = findHeaderKey(headers, ["Unit Price", "Price"]);
      const catKey = findHeaderKey(headers, ["Category"]);
      const typeKey = findHeaderKey(headers, ["Price Type"]);
      const packKey = findHeaderKey(headers, ["Pack Qty"]);
      if (!nameKey || !priceKey) {
        setError("Expected columns: Product Name and Unit Price (SKU, Category, and Price Type optional).");
        return;
      }
      const parsed = json.map(r => {
        const sku = skuKey && r[skuKey] != null ? String(r[skuKey]).trim() : "";
        const name = r[nameKey] == null ? "" : String(r[nameKey]).trim();
        const rawPrice = r[priceKey];
        const unitPrice = rawPrice === "" || rawPrice == null ? NaN : Number(rawPrice);
        const category = catKey && r[catKey] ? String(r[catKey]).trim() : "Other";
        const priceType = typeKey && r[typeKey] ? String(r[typeKey]).trim() : "Per Unit";
        const packQty = packKey && r[packKey] ? Number(r[packKey]) : 1;
        const valid = !!name && Number.isFinite(unitPrice) && unitPrice >= 0;
        const isDuplicate = !!sku && existingSkus.has(sku);
        return {
          id: uid(), sku, name, category, unitPrice, priceType, packQty,
          valid, isDuplicate, include: valid && !isDuplicate,
        };
      });
      setBulkRows(parsed);
    } catch (e) {
      setError("Couldn't read this file. Make sure it's a valid .xlsx or .csv file.");
    }
  };

  const handleInvoiceFile = async (file) => {
    setError(""); setFileName(file.name); setSavedCount(null); setInvoiceProgress("Reading invoice...");
    try {
      let text = "";
      if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
        const pdfjs = await import("pdfjs-dist/build/pdf.mjs");
        const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), disableWorker: true }).promise;
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
          const page = await pdf.getPage(pageNumber);
          const content = await page.getTextContent();
          text += content.items.map(item => item.str).join(" ") + "\n";
        }
      } else {
        const { createWorker } = await import("tesseract.js");
        const worker = await createWorker("eng", 1, { logger: message => { if (message.status) setInvoiceProgress(`${message.status} ${Math.round((message.progress || 0) * 100)}%`); } });
        const result = await worker.recognize(file);
        text = result.data.text;
        await worker.terminate();
      }
      const parsed = parseInvoiceText(text, existingSkus);
      if (!parsed.length) throw new Error("No product-like rows were found.");
      setInvoiceRows(parsed); setInvoiceProgress(`Found ${parsed.length} possible product row${parsed.length === 1 ? "" : "s"}. Review before importing.`);
    } catch (e) {
      console.error(e);
      setError("Could not extract product rows. Try a clearer image or a PDF with selectable text.");
      setInvoiceProgress("");
    }
  };

  const updateInvoiceRow = (id, patch) => setInvoiceRows(rs => rs.map(r => r.id === id ? { ...r, ...patch } : r));
  const invoiceIncludedCount = invoiceRows.filter(r => r.include && r.name.trim() && !isNaN(Number(r.unitPrice))).length;
  const handleImportInvoice = () => {
    const toAdd = invoiceRows.filter(r => r.include && r.name.trim() && !isNaN(Number(r.unitPrice))).map(r => ({
      sku: r.sku.trim() || `INVOICE-${uid().toUpperCase()}`, name: r.name.trim(), category: r.category || "Other",
      unitPrice: Math.round(Number(r.unitPrice) * 100) / 100, priceType: r.priceType || "Per Unit", packQty: Number(r.packQty) || 1
    }));
    if (!toAdd.length) return;
    onAddProducts(toAdd); setSavedCount(toAdd.length); setInvoiceRows([]); setInvoiceProgress(""); setFileName("");
  };

  const toggleInclude = (id) => setBulkRows(rs => rs.map(r => r.id === id ? { ...r, include: !r.include } : r));
  const includedCount = bulkRows.filter(r => r.include).length;

  const handleImportBulk = () => {
    const toAdd = bulkRows.filter(r => r.include).map(r => ({
      sku: r.sku || `INVOICE-${uid().toUpperCase()}`, name: r.name, category: r.category, unitPrice: Math.round(r.unitPrice * 100) / 100,
      priceType: r.priceType, packQty: r.packQty || 1
    }));
    if (toAdd.length === 0) return;
    onAddProducts(toAdd);
    setSavedCount(toAdd.length);
    setBulkRows([]); setFileName("");
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div style={{ padding: "28px 36px", maxWidth: 980 }}>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Add Products</h1>
      <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 20 }}>Add new items to the product database — one at a time, several at once, or by uploading a sheet.</p>

      <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
        {[["manual", "Add Manually", Pencil], ["bulk", "Bulk Upload", Upload], ["invoice", "Invoice PDF / Image", FileSpreadsheet]].map(([id, label, Icon]) => (
          <button key={id} onClick={() => { setMode(id); setSavedCount(null); }}
            style={{
              display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 8,
              border: mode === id ? "1px solid #171a21" : "1px solid #e8e9ee",
              background: mode === id ? "#171a21" : "#fff", color: mode === id ? "#fff" : "#3a3f4c",
              fontSize: 12.5, fontWeight: 600, cursor: "pointer", fontFamily: "'Inter', sans-serif"
            }}>
            <Icon size={13} /> {label}
          </button>
        ))}
      </div>

      {savedCount !== null && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#effaf3", border: "1px solid #cdeddb", borderRadius: 8, padding: "10px 14px", marginBottom: 18, fontSize: 12.5, color: "#166534" }}>
          <CheckCircle2 size={15} /> Added {savedCount} product{savedCount !== 1 ? "s" : ""} to the database.
        </div>
      )}

      {mode === "manual" && (
        <>
          <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", marginBottom: 14 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: "#f8f8fa", textAlign: "left" }}>
                  {["SKU", "Product Name", "Category", "Unit Price", "Price Type", "Pack Qty", ""].map(h => (
                    <th key={h} style={{ padding: "8px 10px", fontWeight: 600, color: "#6b7080", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.03em" }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {manualRows.map(r => {
                  const dup = r.sku.trim() && existingSkus.has(r.sku.trim());
                  return (
                    <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4" }}>
                      <td style={{ padding: "6px 8px" }}>
                        <input value={r.sku} onChange={e => updateManualRow(r.id, { sku: e.target.value })} placeholder="SKU"
                          style={{ width: 100, padding: "5px 7px", border: dup ? "1px solid #e8a33d" : "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }} />
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <input value={r.name} onChange={e => updateManualRow(r.id, { name: e.target.value })} placeholder="Product name"
                          style={{ width: 190, padding: "5px 7px", border: "1px solid #ddd", borderRadius: 5, fontSize: 12.5 }} />
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <select value={r.category} onChange={e => updateManualRow(r.id, { category: e.target.value })}
                          style={{ padding: "5px 6px", border: "1px solid #ddd", borderRadius: 5, fontSize: 12 }}>
                          {CATEGORY_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <input type="number" step="0.01" value={r.unitPrice} onChange={e => updateManualRow(r.id, { unitPrice: e.target.value })} placeholder="0.00"
                          style={{ width: 72, padding: "5px 7px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }} />
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <select value={r.priceType} onChange={e => updateManualRow(r.id, { priceType: e.target.value })}
                          style={{ padding: "5px 6px", border: "1px solid #ddd", borderRadius: 5, fontSize: 12 }}>
                          <option value="Per Unit">Per Unit</option>
                          <option value="Pack Price">Pack Price</option>
                        </select>
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <input type="number" value={r.packQty} onChange={e => updateManualRow(r.id, { packQty: e.target.value })}
                          style={{ width: 52, padding: "5px 7px", border: "1px solid #ddd", borderRadius: 5, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }} />
                      </td>
                      <td style={{ padding: "6px 8px" }}>
                        <Trash2 size={14} color="#c3c7d1" style={{ cursor: "pointer" }} onClick={() => removeManualRow(r.id)} />
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <button onClick={addManualRow} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 5, padding: "7px 12px" }}>
              <PlusCircle size={13} /> Add row
            </button>
            <button
              disabled={manualValidRows.length === 0}
              onClick={handleSaveManual}
              style={{
                background: manualValidRows.length ? "#171a21" : "#d8dbe3", color: "#fff", border: "none",
                padding: "10px 20px", borderRadius: 9, fontSize: 13.5, fontWeight: 600,
                cursor: manualValidRows.length ? "pointer" : "not-allowed", fontFamily: "'Inter', sans-serif"
              }}
            >
              Save {manualValidRows.length > 0 ? manualValidRows.length : ""} Product{manualValidRows.length !== 1 ? "s" : ""}
            </button>
          </div>
          <div style={{ fontSize: 11.5, color: "#aaa", marginTop: 8 }}>Rows with an amber SKU border already exist in the database and will be treated as a new duplicate entry — give them a unique SKU instead.</div>
        </>
      )}

      {mode === "bulk" && (
        <>
          {bulkRows.length === 0 ? (
            <div
              onDragOver={e => e.preventDefault()}
              onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) handleBulkFile(e.dataTransfer.files[0]); }}
              onClick={() => fileInputRef.current?.click()}
              style={{ border: "2px dashed #d8dbe3", borderRadius: 12, padding: "44px 20px", textAlign: "center", cursor: "pointer", background: "#fbfbfc" }}
            >
              <Upload size={26} color="#a3a8b8" style={{ marginBottom: 10 }} />
              <div style={{ fontFamily: "'Inter', sans-serif", fontWeight: 600, fontSize: 14.5, color: "#3a3f4c" }}>Drop a product sheet here, or click to browse</div>
              <div style={{ fontSize: 12, color: "#9aa0ae", marginTop: 5 }}>.xlsx or .csv — columns: SKU, Product Name, Unit Price (Category and Price Type optional)</div>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
                onChange={e => e.target.files[0] && handleBulkFile(e.target.files[0])} />
            </div>
          ) : (
            <>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: "#666" }}>
                  <FileSpreadsheet size={14} /> {fileName} — {bulkRows.length} rows, {includedCount} selected to import
                </div>
                <button onClick={() => { setBulkRows([]); setFileName(""); }} style={ghostBtn}>Start over</button>
              </div>
              <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", marginBottom: 16 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead>
                    <tr style={{ background: "#f8f8fa", textAlign: "left" }}>
                      {["Include", "SKU", "Product Name", "Category", "Unit Price", "Status"].map(h => (
                        <th key={h} style={{ padding: "8px 10px", fontWeight: 600, color: "#6b7080", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.03em" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {bulkRows.map(r => (
                      <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4" }}>
                        <td style={{ padding: "7px 10px" }}>
                          <input type="checkbox" checked={r.include} disabled={!r.valid} onChange={() => toggleInclude(r.id)} />
                        </td>
                        <td style={{ padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace", color: "#666" }}>{r.sku || "—"}</td>
                        <td style={{ padding: "7px 10px", fontWeight: 500 }}>{r.name || "—"}</td>
                        <td style={{ padding: "7px 10px" }}>{r.category ? <CategoryBadge category={r.category} /> : "—"}</td>
                        <td style={{ padding: "7px 10px", fontFamily: "'IBM Plex Mono', monospace" }}>{isNaN(r.unitPrice) ? "—" : money(r.unitPrice)}</td>
                        <td style={{ padding: "7px 10px" }}>
                          {!r.valid ? <StatusPill status="unmatched" /> : r.isDuplicate ? <StatusPill status="ambiguous" /> : <StatusPill status="matched" />}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <button
                disabled={includedCount === 0}
                onClick={handleImportBulk}
                style={{
                  background: includedCount ? "#171a21" : "#d8dbe3", color: "#fff", border: "none",
                  padding: "10px 20px", borderRadius: 9, fontSize: 13.5, fontWeight: 600,
                  cursor: includedCount ? "pointer" : "not-allowed", fontFamily: "'Inter', sans-serif"
                }}
              >
                Import {includedCount} Product{includedCount !== 1 ? "s" : ""}
              </button>
            </>
          )}
        </>
      )}
      {mode === "invoice" && (
        <>
          {invoiceRows.length === 0 ? (
            <div onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) handleInvoiceFile(e.dataTransfer.files[0]); }} onClick={() => fileInputRef.current?.click()}
              style={{ border: "2px dashed #d8dbe3", borderRadius: 12, padding: "44px 20px", textAlign: "center", cursor: "pointer", background: "#fbfbfc" }}>
              <FileSpreadsheet size={26} color="#a3a8b8" style={{ marginBottom: 10 }} />
              <div style={{ fontWeight: 600, fontSize: 14.5, color: "#3a3f4c" }}>Drop a supplier invoice here, or click to browse</div>
              <div style={{ fontSize: 12, color: "#9aa0ae", marginTop: 5 }}>PDF, JPG, or PNG — extracted rows must be reviewed before saving</div>
              <input ref={fileInputRef} type="file" accept=".pdf,.jpg,.jpeg,.png,application/pdf,image/jpeg,image/png" style={{ display: "none" }} onChange={e => e.target.files[0] && handleInvoiceFile(e.target.files[0])} />
            </div>
          ) : (
            <>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                <div style={{ fontSize: 12.5, color: "#666" }}>{fileName} — {invoiceRows.length} possible rows</div>
                <button onClick={() => { setInvoiceRows([]); setFileName(""); setInvoiceProgress(""); }} style={ghostBtn}>Start over</button>
              </div>
              <div style={{ background: "#fff8e1", border: "1px solid #f2df9c", borderRadius: 8, padding: "10px 12px", marginBottom: 12, color: "#765b11", fontSize: 12.5 }}>
                Extraction can misread names, SKUs, prices, or quantities. Check every selected row before importing.
              </div>
              {invoiceProgress && <div style={{ color: "#767c8c", fontSize: 12, marginBottom: 10 }}>{invoiceProgress}</div>}
              <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "auto", marginBottom: 16 }}>
                <table style={{ width: "100%", minWidth: 820, borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead><tr style={{ background: "#f8f8fa", textAlign: "left" }}>{["Include", "SKU", "Product Name", "Category", "Unit Price", "Pack Qty", "Status"].map(h => <th key={h} style={{ padding: "8px 10px", color: "#6b7080", fontSize: 11, textTransform: "uppercase" }}>{h}</th>)}</tr></thead>
                  <tbody>{invoiceRows.map(r => <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4" }}>
                    <td style={{ padding: 7 }}><input type="checkbox" checked={r.include} disabled={!r.valid} onChange={() => updateInvoiceRow(r.id, { include: !r.include })} /></td>
                    <td style={{ padding: 7 }}><input value={r.sku} onChange={e => updateInvoiceRow(r.id, { sku: e.target.value })} style={{ width: 125, padding: 5 }} /></td>
                    <td style={{ padding: 7 }}><input value={r.name} onChange={e => updateInvoiceRow(r.id, { name: e.target.value, valid: !!e.target.value && !isNaN(Number(r.unitPrice)) })} style={{ width: 240, padding: 5 }} /></td>
                    <td style={{ padding: 7 }}><select value={r.category} onChange={e => updateInvoiceRow(r.id, { category: e.target.value })} style={{ padding: 5 }}>{CATEGORY_OPTIONS.map(c => <option key={c}>{c}</option>)}</select></td>
                    <td style={{ padding: 7 }}><input type="number" step="0.01" value={isNaN(r.unitPrice) ? "" : r.unitPrice} onChange={e => updateInvoiceRow(r.id, { unitPrice: e.target.value, valid: !!r.name && e.target.value !== "" && !isNaN(Number(e.target.value)) })} style={{ width: 90, padding: 5 }} /></td>
                    <td style={{ padding: 7 }}><input type="number" min="1" value={r.packQty} onChange={e => updateInvoiceRow(r.id, { packQty: e.target.value })} style={{ width: 60, padding: 5 }} /></td>
                    <td style={{ padding: 7 }}>{!r.valid ? <StatusPill status="unmatched" /> : r.isDuplicate ? <StatusPill status="ambiguous" /> : <StatusPill status="fuzzy" />}</td>
                  </tr>)}</tbody>
                </table>
              </div>
              <button disabled={!invoiceIncludedCount} onClick={handleImportInvoice} style={{ background: invoiceIncludedCount ? "#171a21" : "#d8dbe3", color: "#fff", border: "none", padding: "10px 20px", borderRadius: 9, fontSize: 13.5, fontWeight: 600, cursor: invoiceIncludedCount ? "pointer" : "not-allowed" }}>
                Add {invoiceIncludedCount || ""} Approved Product{invoiceIncludedCount !== 1 ? "s" : ""}
              </button>
            </>
          )}
        </>
      )}
      {error && <div style={{ color: "#a5202a", fontSize: 12.5, marginTop: 12 }}>{error}</div>}
    </div>
  );
}

function InvoiceToExcelView() {
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [progress, setProgress] = useState("");
  const [error, setError] = useState("");
  const [model, setModel] = useState("prebuilt-invoice");
  const inputRef = useRef(null);

  const readInvoice = async (file) => {
    setError(""); setRows([]); setFileName(file.name); setProgress("Reading invoice...");
    try {
      const buffer = await file.arrayBuffer();
      let binary = "";
      const bytes = new Uint8Array(buffer);
      for (let offset = 0; offset < bytes.length; offset += 0x8000) binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
      setProgress("Analyzing invoice with Azure...");
      const response = await fetch("/api/invoice-scan", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileName: file.name, mimeType: file.type, base64: btoa(binary), model }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Azure could not analyze this file.");
      const parsed = (result.items || []).map(r => ({ id: uid(), name: String(r.name || ""), price: Number(r.price) })).filter(r => r.name && Number.isFinite(r.price));
      if (!parsed.length) throw new Error("No product rows found");
      setRows(parsed); setProgress(`Azure found ${parsed.length} possible product row${parsed.length === 1 ? "" : "s"}. Review the names and prices before exporting.`);
    } catch (e) {
      console.error(e); setError(e.message || "Could not extract product rows. Try a clearer image or PDF."); setProgress("");
    }
  };
  const updateRow = (id, patch) => setRows(rs => rs.map(r => r.id === id ? { ...r, ...patch } : r));
  const exportRows = rows.filter(r => r.name.trim() && Number.isFinite(Number(r.price)));
  const download = () => exportToExcel("invoice_products.xlsx", [{ name: "Products", rows: exportRows.map(r => ({ "Product Name": r.name.trim(), "Unit Price": Number(r.price) })) }]);

  return <div style={{ padding: "28px 36px", maxWidth: 980 }}>
    <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, margin: 0, color: "#171a21" }}>Invoice to Excel</h1>
    <p style={{ color: "#767c8c", fontSize: 13.5, marginTop: 5, marginBottom: 20 }}>Upload an invoice photo or PDF, review the extracted product names and prices, then download a sheet for Add Products.</p>
    <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
      {[ ["prebuilt-invoice", "Invoice model"], ["prebuilt-receipt", "Receipt model"] ].map(([id, label]) => <button key={id} onClick={() => setModel(id)} style={{ padding: "7px 12px", borderRadius: 7, border: model === id ? "1px solid #171a21" : "1px solid #e8e9ee", background: model === id ? "#171a21" : "#fff", color: model === id ? "#fff" : "#3a3f4c", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{label}</button>)}
    </div>
    {!rows.length ? <div onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); if (e.dataTransfer.files[0]) readInvoice(e.dataTransfer.files[0]); }} onClick={() => inputRef.current?.click()}
      style={{ border: "2px dashed #d8dbe3", borderRadius: 12, padding: "44px 20px", textAlign: "center", cursor: "pointer", background: "#fbfbfc" }}>
      <ScanLine size={26} color="#a3a8b8" style={{ marginBottom: 10 }} />
      <div style={{ fontWeight: 600, fontSize: 14.5, color: "#3a3f4c" }}>Drop an invoice image here, or click to browse</div>
      <div style={{ fontSize: 12, color: "#9aa0ae", marginTop: 5 }}>PDF, JPG, or PNG — up to 4 MB. Scanned securely using the selected Azure model.</div>
      <input ref={inputRef} type="file" accept=".pdf,.jpg,.jpeg,.png,application/pdf,image/jpeg,image/png" style={{ display: "none" }} onChange={e => e.target.files[0] && readInvoice(e.target.files[0])} />
    </div> : <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 12.5, color: "#666" }}>{fileName} — {rows.length} possible rows</div>
        <button onClick={() => { setRows([]); setFileName(""); setProgress(""); if (inputRef.current) inputRef.current.value = ""; }} style={ghostBtn}>Start over</button>
      </div>
      <div style={{ background: "#fff8e1", border: "1px solid #f2df9c", borderRadius: 8, padding: "10px 12px", marginBottom: 12, color: "#765b11", fontSize: 12.5 }}>OCR can misread text and prices. Verify every row before downloading.</div>
      <div style={{ color: "#767c8c", fontSize: 12, marginBottom: 10 }}>{progress}</div>
      <div style={{ border: "1px solid #e8e9ee", borderRadius: 10, overflow: "hidden", marginBottom: 16 }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
        <thead><tr style={{ background: "#f8f8fa", textAlign: "left" }}>{["Product Name", "Price"].map(h => <th key={h} style={{ padding: "8px 10px", color: "#6b7080", fontSize: 11, textTransform: "uppercase" }}>{h}</th>)}</tr></thead>
        <tbody>{rows.map(r => <tr key={r.id} style={{ borderTop: "1px solid #f0f1f4" }}>
          <td style={{ padding: 7 }}><input value={r.name} onChange={e => updateRow(r.id, { name: e.target.value })} style={{ width: "90%", padding: 6 }} /></td>
          <td style={{ padding: 7 }}><input type="number" min="0" step="0.01" value={Number.isFinite(Number(r.price)) ? r.price : ""} onChange={e => updateRow(r.id, { price: e.target.value })} style={{ width: 120, padding: 6 }} /></td>
        </tr>)}</tbody>
      </table></div>
      <button disabled={!exportRows.length} onClick={download} style={{ background: exportRows.length ? "#171a21" : "#d8dbe3", color: "#fff", border: "none", padding: "10px 18px", borderRadius: 9, fontSize: 13, fontWeight: 600, cursor: exportRows.length ? "pointer" : "not-allowed", display: "inline-flex", alignItems: "center", gap: 7 }}><Download size={15} /> Download Excel for Add Products</button>
    </>}
    {progress && !rows.length && <div style={{ color: "#767c8c", fontSize: 12, marginTop: 10 }}>{progress}</div>}
    {error && <div style={{ color: "#a5202a", fontSize: 12.5, marginTop: 12 }}>{error}</div>}
  </div>;
}

export default function StoreTransferApp() {
  const [view, setView] = useState("new");
  const [stores] = useState(DEFAULT_STORES);
  const [transfers, setTransfers] = useState([]);
  const [products, setProducts] = useState(PRODUCTS_SEED);
  const [loading, setLoading] = useState(true);
  const [transferToEdit, setTransferToEdit] = useState(null);

  useEffect(() => {
    Promise.all([loadTransfers(), loadProducts()]).then(([t, p]) => {
      setTransfers(t);
      if (Array.isArray(p) && p.length) {
        setProducts(p);
      } else {
        setProducts(PRODUCTS_SEED);
        saveProducts(PRODUCTS_SEED);
      }
      setLoading(false);
    });
  }, []);

  const handleCreateTransfer = useCallback((t) => {
    setTransfers(prev => {
      const next = prev.some(existing => existing.id === t.id)
        ? prev.map(existing => existing.id === t.id ? t : existing)
        : [...prev, t];
      saveTransfers(next);
      return next;
    });
    setTransferToEdit(null);
  }, []);

  const handleDeleteTransfer = useCallback((id) => {
    setTransfers(prev => {
      const next = prev.filter(t => t.id !== id);
      saveTransfers(next);
      return next;
    });
  }, []);

  const handleAddProducts = useCallback((newProducts) => {
    setProducts(prev => {
      const bySku = new Map(prev.map(p => [p.sku, p]));
      newProducts.forEach(p => bySku.set(p.sku, p));
      const next = Array.from(bySku.values());
      saveProducts(next);
      return next;
    });
  }, []);

  const handleUpdateProduct = useCallback((index, patch) => {
    setProducts(prev => {
      if (index < 0 || index >= prev.length) return prev;
      const next = prev.map((product, i) => i === index ? { ...product, ...patch } : product);
      saveProducts(next);
      return next;
    });
  }, []);

  const handleCreateMultipleTransfers = useCallback((newTransfers) => {
    setTransfers(prev => {
      const next = [...prev, ...newTransfers];
      saveTransfers(next);
      return next;
    });
  }, []);

  const handleMarkTransfers = useCallback((markMap) => {
    // markMap: { transferId: [invoiceId, ...], ... }
    setTransfers(prev => {
      const next = prev.map(t => {
        if (markMap[t.id]) {
          const existing = Array.isArray(t.reconciledWith) ? t.reconciledWith : [];
          const merged = Array.from(new Set([...existing, ...markMap[t.id]]));
          return { ...t, reconciledWith: merged };
        }
        return t;
      });
      saveTransfers(next);
      return next;
    });
  }, []);

  const handleEditTransfer = useCallback((transfer) => {
    setTransferToEdit(transfer);
    setView("new");
  }, []);

  const handleViewChange = useCallback((nextView) => {
    if (nextView === "new") setTransferToEdit(null);
    setView(nextView);
  }, []);

  return (
    <div style={{ display: "flex", height: "100%", minHeight: 640, background: "#fff", fontFamily: "'Inter', sans-serif", color: "#171a21" }}>
      <style>{FONT_IMPORT}{`
        * { box-sizing: border-box; }
        input:focus, select:focus { outline: 2px solid #e8a33d33; border-color: #e8a33d; }
        @media print {
          .no-print, aside, [data-sidebar] { display: none !important; }
        }
      `}</style>
      <Sidebar view={view} setView={handleViewChange} stores={stores} productCount={products.length} />
      <div style={{ flex: 1, overflowY: "auto" }}>
        {view === "new" && <NewTransferView key={transferToEdit?.id || "new"} stores={stores} products={products} onCreateTransfer={handleCreateTransfer} initialTransfer={transferToEdit} />}
        {view === "reconcile" && <ReconcileView transfers={transfers} products={products} stores={stores} onCreateTransfers={handleCreateMultipleTransfers} onMarkTransfers={handleMarkTransfers} />}
        {view === "history" && <HistoryView transfers={transfers} loading={loading} onEdit={handleEditTransfer} onDelete={handleDeleteTransfer} />}
        {view === "dashboard" && <DashboardView transfers={transfers} stores={stores} />}
        {view === "products" && <ProductsView products={products} onUpdateProduct={handleUpdateProduct} />}
        {view === "addproducts" && <AddProductsView products={products} onAddProducts={handleAddProducts} />}
        {view === "invoiceexcel" && <InvoiceToExcelView />}
      </div>
    </div>
  );
}
