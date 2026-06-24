#!/usr/bin/env python3
# Writes modul_v2_clean.js with all emoji properly unicode-escaped

JS = r"""
'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak
} = require('docx');
const fs = require('fs');

// ── Palette ─────────────────────────────────────────────────────────────
const C = {
  teal:"00796B",tealD:"004D40",tealL:"B2DFDB",tealXL:"E0F2F1",
  orange:"E65100",orangeM:"F57C00",orangeL:"FFE0B2",orangeXL:"FFF3E0",
  purple:"6A1B9A",purpleM:"8E24AA",purpleL:"E1BEE7",purpleXL:"F3E5F5",
  blue:"1565C0",blueM:"1976D2",blueL:"BBDEFB",blueXL:"E3F2FD",
  green:"2E7D32",greenM:"388E3C",greenL:"C8E6C9",greenXL:"E8F5E9",
  red:"B71C1C",redM:"C62828",redL:"FFCDD2",redXL:"FFEBEE",
  pink:"AD1457",pinkL:"FCE4EC",
  amber:"F57F17",amberL:"FFF8E1",
  white:"FFFFFF",black:"212121",
  gray:"37474F",grayM:"546E7A",grayL:"CFD8DC",grayXL:"ECEFF1",
  yellow:"F9A825",yellowL:"FFFDE7",
};

// ── Helpers ──────────────────────────────────────────────────────────────
const bdr  = (color,size=6) => ({style:BorderStyle.SINGLE,size,color});
const bdrT = (color,size=4) => ({top:bdr(color,size),bottom:bdr(color,size),left:bdr(color,size),right:bdr(color,size)});
const bdrN = ()             => ({top:bdr(C.white,1),bottom:bdr(C.white,1),left:bdr(C.white,1),right:bdr(C.white,1)});
const shade= (fill)         => ({fill,type:ShadingType.CLEAR});

const run = (text,{bold=false,italic=false,size=22,color=C.black,font="Arial"}={}) =>
  new TextRun({text,bold,italic,size,color,font});

const p = (children,{align=AlignmentType.LEFT,before=60,after=60,indent={}}={}) =>
  new Paragraph({
    children: Array.isArray(children)?children:[children],
    alignment:align, spacing:{before,after}, indent
  });

const pRun = (text,opts={},pOpts={}) => p([run(text,opts)],pOpts);

const spacer = (h=100) => new Paragraph({children:[run("")],spacing:{before:h,after:0}});

const mkCell = (children,{w,fill=C.white,borders,vAlign=VerticalAlign.TOP,span}={}) =>
  new TableCell({
    children: Array.isArray(children)?children:[children],
    width:{size:w,type:WidthType.DXA},
    shading:shade(fill),
    borders: borders||bdrT(C.grayL,3),
    margins:{top:90,bottom:90,left:150,right:150},
    verticalAlign:vAlign,
    ...(span?{columnSpan:span}:{}),
  });

const mkHdr = (text,{w,fill=C.teal,color=C.white,size=20,span}={}) =>
  new TableCell({
    children:[p([run(text,{bold:true,size,color})],{align:AlignmentType.CENTER,before:70,after:70})],
    width:{size:w,type:WidthType.DXA},
    shading:shade(fill),
    borders:bdrT(fill,6),
    margins:{top:70,bottom:70,left:120,right:120},
    verticalAlign:VerticalAlign.CENTER,
    ...(span?{columnSpan:span}:{}),
  });

const banner = (text,{fill=C.teal,textColor=C.white,size=26}={}) =>
  new Table({
    width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
    rows:[new TableRow({children:[new TableCell({
      children:[p([run(text,{bold:true,size,color:textColor})],{before:110,after:110})],
      width:{size:9360,type:WidthType.DXA},shading:shade(fill),borders:bdrN(),
      margins:{top:70,bottom:70,left:180,right:180},
    })]})],
  });

const subBanner = (title,{fill=C.tealXL,border=C.teal}={}) =>
  new Paragraph({
    children:[run(title,{bold:true,size:23,color:border})],
    shading:shade(fill),spacing:{before:110,after:70},indent:{left:160},
    border:{left:{style:BorderStyle.SINGLE,size:16,color:border}},
  });

const bullet = (text,{color=C.black,bold=false,ref="bul"}={}) =>
  new Paragraph({
    numbering:{reference:ref,level:0},
    children:[run(text,{color,bold})],spacing:{before:55,after:55},
  });

const num = (text,ref="num1",{bold=false,color=C.black}={}) =>
  new Paragraph({
    numbering:{reference:ref,level:0},
    children:[run(text,{bold,color})],spacing:{before:65,after:65},
  });

const callout = (titlePara, bodyItems, {fill=C.blueXL,accent=C.blue,w=9360}={}) =>
  new Table({
    width:{size:w,type:WidthType.DXA},columnWidths:[w],
    rows:[
      new TableRow({children:[new TableCell({
        children:[titlePara],width:{size:w,type:WidthType.DXA},
        shading:shade(accent),borders:bdrT(accent,6),
        margins:{top:80,bottom:80,left:180,right:180},
      })]}),
      ...bodyItems.map(r=>new TableRow({children:[new TableCell({
        children:Array.isArray(r)?r:[r],width:{size:w,type:WidthType.DXA},
        shading:shade(fill),borders:bdrT(accent,4),
        margins:{top:70,bottom:70,left:190,right:190},
      })]})),
    ],
  });

// ── Numbering configs ─────────────────────────────────────────────────────
const numConfigs = [
  { reference:"bul", levels:[
    {level:0,format:LevelFormat.BULLET,text:"\u2022",
     style:{paragraph:{indent:{left:580,hanging:280}}}},
    {level:1,format:LevelFormat.BULLET,text:"\u25E6",
     style:{paragraph:{indent:{left:880,hanging:280}}}},
  ]},
];
for(let n=1;n<=12;n++){
  numConfigs.push({reference:`num${n}`,levels:[
    {level:0,format:LevelFormat.DECIMAL,text:`%1.`,
     style:{paragraph:{indent:{left:620,hanging:300}}}},
  ]});
}

// ── Document ──────────────────────────────────────────────────────────────
const doc = new Document({
  numbering:{config:numConfigs},
  styles:{
    default:{document:{run:{font:"Arial",size:22,color:C.black}}},
    paragraphStyles:[
      {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,
       run:{size:32,bold:true,font:"Arial",color:C.white},
       paragraph:{spacing:{before:240,after:120},outlineLevel:0}},
      {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
       run:{size:26,bold:true,font:"Arial",color:C.orange},
       paragraph:{spacing:{before:180,after:100},outlineLevel:1}},
    ]
  },
  sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:1000,right:1000,bottom:1000,left:1000}}},
    headers:{default:new Header({children:[
      new Paragraph({
        children:[
          run("MODUL AJAR PJOK \u2022 KELAS VIII \u2022 AKTIVITAS GERAK BERIRAMA",{bold:true,size:17,color:C.teal}),
          run("   |   ",{size:17,color:C.grayL}),
          run("Deep Learning & Pembelajaran Berdiferensiasi",{size:17,color:C.grayM,italic:true}),
        ],
        border:{bottom:{style:BorderStyle.SINGLE,size:8,color:C.teal,space:4}},
        spacing:{before:0,after:100},
      }),
    ]})},
    footers:{default:new Footer({children:[
      new Paragraph({
        children:[
          run("Pendidikan Jasmani, Olahraga & Kesehatan  \u2022  SMP Kelas VIII  \u2022  Fase D",{size:17,color:C.grayM}),
          run("     Hal. ",{size:17,color:C.grayM}),
          new TextRun({children:[PageNumber.CURRENT],font:"Arial",size:17,color:C.teal,bold:true}),
        ],
        border:{top:{style:BorderStyle.SINGLE,size:8,color:C.teal,space:4}},
        spacing:{before:100,after:0},
      }),
    ]})},

    children:[

      // ════════════════════════════════════════════════════════════
      // COVER
      // ════════════════════════════════════════════════════════════

      // Gradient bar
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3120,3120,3120],
        rows:[new TableRow({children:[
          new TableCell({children:[p([run("",{size:4})],{before:0,after:0})],width:{size:3120,type:WidthType.DXA},shading:shade(C.tealD),borders:bdrN(),margins:{top:0,bottom:0,left:0,right:0}}),
          new TableCell({children:[p([run("",{size:4})],{before:0,after:0})],width:{size:3120,type:WidthType.DXA},shading:shade(C.teal),borders:bdrN(),margins:{top:0,bottom:0,left:0,right:0}}),
          new TableCell({children:[p([run("",{size:4})],{before:0,after:0})],width:{size:3120,type:WidthType.DXA},shading:shade(C.orange),borders:bdrN(),margins:{top:0,bottom:0,left:0,right:0}}),
        ]})]
      }),

      // Title block
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
        rows:[new TableRow({children:[new TableCell({
          children:[
            spacer(60),
            p([run("MODUL AJAR",{bold:true,size:22,color:C.tealXL})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("PENDIDIKAN JASMANI, OLAHRAGA & KESEHATAN",{size:20,color:C.white})],{align:AlignmentType.CENTER,before:0,after:30}),
            p([run("\u2605  AKTIVITAS GERAK BERIRAMA  \u2605",{bold:true,size:42,color:C.white})],{align:AlignmentType.CENTER,before:0,after:30}),
            p([run("Kelas VIII  \u2014  Fase D  \u2014  Kurikulum Merdeka",{size:21,color:C.amberL,italic:true})],{align:AlignmentType.CENTER,before:0,after:60}),
            spacer(40),
          ],
          width:{size:9360,type:WidthType.DXA},shading:shade(C.teal),borders:bdrN(),
          margins:{top:60,bottom:60,left:240,right:240},
        })]})],
      }),

      // 3 Pillar cards
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3120,3120,3120],
        rows:[new TableRow({children:[
          new TableCell({children:[
            p([run("\u{1F9D8}",{font:"Segoe UI Emoji",size:36})],{align:AlignmentType.CENTER,before:80,after:20}),
            p([run("MINDFUL",{bold:true,size:24,color:C.teal})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Kesadaran Penuh",{size:18,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Fokus pada gerak, napas, irama & sinkronisasi diri",{size:18,color:C.black})],{align:AlignmentType.CENTER,before:0,after:80}),
          ],width:{size:3120,type:WidthType.DXA},shading:shade(C.tealXL),borders:bdrT(C.teal,8),margins:{top:60,bottom:60,left:120,right:120}}),
          new TableCell({children:[
            p([run("\u{1F31F}",{font:"Segoe UI Emoji",size:36})],{align:AlignmentType.CENTER,before:80,after:20}),
            p([run("MEANINGFUL",{bold:true,size:24,color:C.orange})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Bermakna",{size:18,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Gerak dikaitkan budaya, kebugaran & kehidupan nyata",{size:18,color:C.black})],{align:AlignmentType.CENTER,before:0,after:80}),
          ],width:{size:3120,type:WidthType.DXA},shading:shade(C.orangeXL),borders:bdrT(C.orange,8),margins:{top:60,bottom:60,left:120,right:120}}),
          new TableCell({children:[
            p([run("\u{1F3B6}",{font:"Segoe UI Emoji",size:36})],{align:AlignmentType.CENTER,before:80,after:20}),
            p([run("JOYFUL",{bold:true,size:24,color:C.purple})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Penuh Kegembiraan",{size:18,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Musik pop, tantangan gerak & kreativitas bebas",{size:18,color:C.black})],{align:AlignmentType.CENTER,before:0,after:80}),
          ],width:{size:3120,type:WidthType.DXA},shading:shade(C.purpleXL),borders:bdrT(C.purple,8),margins:{top:60,bottom:60,left:120,right:120}}),
        ]})]
      }),

      spacer(60),

      // Identity table
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3200,6160],
        rows:[
          new TableRow({children:[mkHdr("IDENTITAS MODUL",{w:9360,span:2,fill:C.tealD,size:21})]}),
          ...[
            ["Mata Pelajaran","Pendidikan Jasmani, Olahraga & Kesehatan (PJOK)"],
            ["Kelas / Semester","VIII (Delapan) / Semester 1"],
            ["Materi Pokok","Aktivitas Gerak Berirama"],
            ["Alokasi Waktu","2 Pertemuan \u00D7 2 \u00D7 40 menit (Total 160 Menit)"],
            ["Pendekatan","Deep Learning \u2014 Mindful, Meaningful, Joyful"],
            ["Model Pembelajaran","Pembelajaran Berdiferensiasi (Konten, Proses, Produk)"],
            ["Fase Kurikulum","Fase D \u2014 Kurikulum Merdeka"],
            ["Sarana & Prasarana","Lapangan/Aula, Speaker, Pita, Bola Kecil, Proyektor/TV"],
          ].map(([k,v],i)=>new TableRow({children:[
            mkCell(pRun(k,{bold:true,size:20,color:C.teal}),{w:3200,fill:i%2===0?C.tealXL:C.blueXL}),
            mkCell(pRun(v,{size:20}),{w:6160,fill:i%2===0?C.white:C.grayXL}),
          ]})),
        ],
      }),

      spacer(80),
      pRun("Dibuat oleh: ________________________________     NIP: ____________________________",{size:19,color:C.grayM},{before:0,after:40}),
      pRun("Sekolah: ____________________________________________     Tahun Pelajaran: _________",{size:19,color:C.grayM},{before:0,after:0}),
      spacer(40),

      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 1: TUJUAN PEMBELAJARAN
      // ════════════════════════════════════════════════════════════
      banner("1.  TUJUAN PEMBELAJARAN",{fill:C.teal}),
      spacer(70),
      pRun("Tujuan pembelajaran dirumuskan berdasarkan Capaian Pembelajaran Fase D, mencakup tiga domain:"),
      spacer(60),

      subBanner("\u{1F9E0}  A. Domain KOGNITIF \u2014 Pengetahuan & Pemahaman",{fill:C.blueXL,border:C.blue}),
      num("Menganalisis hubungan antara gerak tubuh, irama musik, tempo, dan koordinasi dalam aktivitas gerak berirama.","num1"),
      num("Menjelaskan konsep keseimbangan, keluwesan, dan ketepatan tempo dalam rangkaian gerak berirama.","num1"),
      num("Mengidentifikasi unsur-unsur gerak berirama dalam kehidupan sehari-hari, seni, budaya, dan olahraga.","num1"),
      num("Membedakan gerak yang berirama dan tidak berirama berdasarkan pengamatan langsung.","num1"),
      spacer(60),

      subBanner("\u{1F3CB}  B. Domain PSIKOMOTOR \u2014 Keterampilan Gerak",{fill:C.orangeXL,border:C.orange}),
      num("Mempraktikkan langkah dasar (march, step touch, grapevine) dan ayunan tangan sesuai irama dengan koordinasi benar.","num2"),
      num("Melakukan rangkaian gerak berirama 2\u00D78 hitungan dengan keluwesan, keseimbangan, dan ketepatan tempo.","num2"),
      num("Menggabungkan gerak tangan, kaki, dan badan secara harmonis sesuai tempo musik yang bervariasi.","num2"),
      num("Menampilkan kreasi gerak berirama berkelompok dengan variasi pola lantai dan ekspresi positif.","num2"),
      spacer(60),

      subBanner("\u2764  C. Domain AFEKTIF \u2014 Sikap & Karakter",{fill:C.purpleXL,border:C.purple}),
      num("Menunjukkan percaya diri saat mempraktikkan gerak berirama di depan guru dan teman sekelas.","num3"),
      num("Menerapkan kerja sama, komunikasi positif, dan saling menghargai dalam aktivitas gerak kelompok.","num3"),
      num("Menampilkan disiplin, tanggung jawab, dan semangat belajar aktif selama seluruh sesi pembelajaran.","num3"),
      num("Merespons umpan balik dari guru dan teman dengan sikap terbuka dan kemauan memperbaiki diri.","num3"),
      spacer(70),

      // Profil Pelajar Pancasila
      callout(
        p([run("\u{1F1EE}\u{1F1E9}  Profil Pelajar Pancasila yang Dikembangkan",{bold:true,size:21,color:C.white})],{before:0,after:0}),
        [new Table({width:{size:9160,type:WidthType.DXA},columnWidths:[2290,2290,2290,2290],
          rows:[new TableRow({children:[
            mkCell([p([run("\u{1F91D} Bergotong Royong",{bold:true,size:20,color:C.teal})],{before:0,after:20}),pRun("Kerja sama & sinkronisasi kelompok",{size:18})],{w:2290,fill:C.tealXL}),
            mkCell([p([run("\u{1F4AA} Mandiri",{bold:true,size:20,color:C.blue})],{before:0,after:20}),pRun("Evaluasi diri & motivasi gerak",{size:18})],{w:2290,fill:C.blueXL}),
            mkCell([p([run("\u{1F3A8} Kreatif",{bold:true,size:20,color:C.purple})],{before:0,after:20}),pRun("Kreasi variasi gerak & pola lantai",{size:18})],{w:2290,fill:C.purpleXL}),
            mkCell([p([run("\u{1F30D} Bernalar Kritis",{bold:true,size:20,color:C.orange})],{before:0,after:20}),pRun("Analisis gerak, tempo & koordinasi",{size:18})],{w:2290,fill:C.orangeXL}),
          ]})]
        })],
        {fill:C.tealXL,accent:C.teal,w:9360}
      ),

      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 2: PENDEKATAN DEEP LEARNING
      // ════════════════════════════════════════════════════════════
      banner("2.  PENDEKATAN DEEP LEARNING",{fill:C.tealD}),
      spacer(70),

      subBanner("\u{1F9D8}  A. MINDFUL \u2014 Kesadaran Penuh dalam Bergerak",{fill:C.tealXL,border:C.teal}),
      pRun("Peserta didik dikembangkan kesadaran penuh (mindfulness) melalui:"),
      bullet("Memperhatikan setiap gerakan tubuh secara sadar: posisi tangan, pijakan kaki, dan orientasi tubuh."),
      bullet("Mengatur napas secara teratur dan sadar selaras dengan ritme musik yang dimainkan."),
      bullet("Mengenali dan membedakan tempo musik: lambat (60 BPM), sedang (90 BPM), dan cepat (120 BPM)."),
      bullet("Merasakan sinkronisasi gerak antaranggota kelompok melalui kepekaan indera dan fokus kolektif."),
      bullet("Berlatih hadir penuh (present): meminimalkan distraksi dan merasakan kegembiraan gerak secara utuh."),
      bullet("Sesi jeda reflektif 30 detik di antara latihan untuk body scan sederhana (memindai perasaan tubuh)."),
      spacer(60),

      subBanner("\u{1F31F}  B. MEANINGFUL \u2014 Pembelajaran yang Bermakna",{fill:C.orangeXL,border:C.orange}),
      pRun("Keterkaitan gerak berirama dengan dunia nyata peserta didik:"),
      spacer(50),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2800,6560],
        rows:[
          new TableRow({children:[mkHdr("Konteks Kehidupan",{w:2800,fill:C.orange}),mkHdr("Relevansi dengan Gerak Berirama",{w:6560,fill:C.orange})]}),
          ...[
            ["\u{1F3B5} Dance Challenge Media Sosial","Gerakan viral TikTok/Reels mengajarkan koordinasi, ekspresi, dan kreativitas gerak yang kontekstual bagi remaja."],
            ["\u{1F3C3} Kebugaran & Kesehatan Tubuh","Gerak berirama membangun daya tahan kardiorespirasi, fleksibilitas sendi, dan koordinasi neuromuskular."],
            ["\u{1F3AD} Seni & Budaya Indonesia","Elemen tari tradisional (Poco-Poco, Sajojo) mencerminkan identitas budaya yang relevan dengan gerak ritmis."],
            ["\u{1F91D} Kerja Sama Tim","Sinkronisasi gerak berkelompok melatih komunikasi non-verbal, kepercayaan, dan tanggung jawab kolektif."],
            ["\u{1F393} Persiapan Porseni/Pentas Seni","Keterampilan gerak berirama menjadi modal untuk penampilan seni di sekolah dan lomba."],
          ].map(([k,v],i)=>new TableRow({children:[
            mkCell(pRun(k,{bold:true,size:20,color:C.orange}),{w:2800,fill:i%2===0?C.orangeXL:C.amberL}),
            mkCell(pRun(v,{size:20}),{w:6560,fill:i%2===0?C.white:C.grayXL}),
          ]})),
        ]
      }),
      spacer(60),

      subBanner("\u{1F3B6}  C. JOYFUL \u2014 Kegembiraan Intrinsik dalam Bergerak",{fill:C.purpleXL,border:C.purple}),
      bullet("Menggunakan playlist musik populer remaja (K-Pop, Pop Indonesia, EDM ringan) yang energik dan familiar."),
      bullet('"Momen Bintang" di akhir sesi: setiap kelompok tampil 30 detik, mendapat tepuk tangan dan poin apresiasi.'),
      bullet("Gamifikasi: sistem poin untuk ketepatan tempo, kekompakan, dan kreativitas gerak kelompok."),
      bullet("Variasi musik di setiap pertemuan agar tidak monoton dan mempertahankan antusiasme peserta."),
      bullet("Ruang kreasi bebas: peserta didik boleh menambahkan gerakan khas diri sendiri dalam kelompok."),
      bullet("Zero judgment zone: semua level kemampuan diapresiasi; kesalahan adalah bagian dari proses belajar."),
      bullet('"Freeze Dance" mini sebagai reward di akhir kegiatan inti untuk menciptakan momen joyful bersama.'),
      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 3: ASESMEN DIAGNOSTIK
      // ════════════════════════════════════════════════════════════
      banner("3.  ASESMEN DIAGNOSTIK AWAL",{fill:C.blue}),
      spacer(70),
      pRun("Guru memetakan kesiapan belajar SEBELUM kegiatan inti dimulai. Hasil digunakan untuk pengelompokan stasiun."),
      spacer(70),

      // Q1 Card
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[700,8660],
        rows:[
          new TableRow({children:[
            new TableCell({children:[p([run("Q1",{bold:true,size:26,color:C.white})],{align:AlignmentType.CENTER,before:80,after:80})],
              width:{size:700,type:WidthType.DXA},shading:shade(C.teal),borders:bdrT(C.teal,6),
              margins:{top:80,bottom:80,left:80,right:80},verticalAlign:VerticalAlign.CENTER}),
            new TableCell({children:[
              p([run("Pertanyaan Pemantik 1 \u2014 Pengalaman Awal",{bold:true,size:20,color:C.teal})],{before:0,after:40}),
              p([run('"Siapa yang pernah ikut senam, tari tradisional, atau dance challenge? Bagaimana rasanya?"',{italic:true,size:21})],{before:0,after:40}),
              p([run("\u25B6 Tujuan: Memetakan pengalaman gerak berirama & kesiapan afektif peserta didik.",{size:18,color:C.grayM})],{before:0,after:0}),
            ],width:{size:8660,type:WidthType.DXA},shading:shade(C.tealXL),borders:bdrT(C.teal,4),margins:{top:100,bottom:100,left:180,right:180}}),
          ]}),
          new TableRow({children:[
            new TableCell({children:[p([run("")],{before:30,after:30})],
              width:{size:9360,type:WidthType.DXA},columnSpan:2,shading:shade(C.white),borders:bdrN(),margins:{top:0,bottom:0,left:0,right:0}}),
          ]}),
          // Q2
          new TableRow({children:[
            new TableCell({children:[p([run("Q2",{bold:true,size:26,color:C.white})],{align:AlignmentType.CENTER,before:80,after:80})],
              width:{size:700,type:WidthType.DXA},shading:shade(C.orange),borders:bdrT(C.orange,6),
              margins:{top:80,bottom:80,left:80,right:80},verticalAlign:VerticalAlign.CENTER}),
            new TableCell({children:[
              p([run("Pertanyaan Pemantik 2 \u2014 Kesadaran Tantangan Gerak",{bold:true,size:20,color:C.orange})],{before:0,after:40}),
              p([run('"Menurutmu, gerakan mana yang paling sulit dilakukan sesuai irama? Mengapa?"',{italic:true,size:21})],{before:0,after:40}),
              p([run("\u25B6 Tujuan: Mengidentifikasi persepsi peserta didik tentang tantangan gerak dan memetakan kesiapan kognitif.",{size:18,color:C.grayM})],{before:0,after:0}),
            ],width:{size:8660,type:WidthType.DXA},shading:shade(C.orangeXL),borders:bdrT(C.orange,4),margins:{top:100,bottom:100,left:180,right:180}}),
          ]}),
          new TableRow({children:[
            new TableCell({children:[p([run("")],{before:30,after:30})],
              width:{size:9360,type:WidthType.DXA},columnSpan:2,shading:shade(C.white),borders:bdrN(),margins:{top:0,bottom:0,left:0,right:0}}),
          ]}),
          // Q3
          new TableRow({children:[
            new TableCell({children:[p([run("Q3",{bold:true,size:26,color:C.white})],{align:AlignmentType.CENTER,before:80,after:80})],
              width:{size:700,type:WidthType.DXA},shading:shade(C.purple),borders:bdrT(C.purple,6),
              margins:{top:80,bottom:80,left:80,right:80},verticalAlign:VerticalAlign.CENTER}),
            new TableCell({children:[
              p([run("Aktivitas Diagnostik 3 \u2014 Observasi Psikomotor Langsung",{bold:true,size:20,color:C.purple})],{before:0,after:40}),
              p([run('"Ikuti 8 hitungan gerakan sederhana yang Bapak/Ibu contohkan! Langkah kaki kanan dulu, ikuti temponya."',{italic:true,size:21})],{before:0,after:40}),
              p([run("\u25B6 Tujuan: Observasi langsung koordinasi gerak, kepekaan tempo, dan level kesiapan psikomotor.",{size:18,color:C.grayM})],{before:0,after:0}),
            ],width:{size:8660,type:WidthType.DXA},shading:shade(C.purpleXL),borders:bdrT(C.purple,4),margins:{top:100,bottom:100,left:180,right:180}}),
          ]}),
        ]
      }),

      spacer(80),
      subBanner("\u{1F4CB}  Panduan Pemetaan Hasil Diagnostik ke Level Stasiun",{fill:C.yellowL,border:C.amber}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1600,3880,3880],
        rows:[
          new TableRow({children:[mkHdr("Level",{w:1600,fill:C.gray}),mkHdr("Ciri-Ciri yang Diamati",{w:3880,fill:C.gray}),mkHdr("Penempatan Stasiun",{w:3880,fill:C.gray})]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F331} PEMULA",{bold:true,size:20,color:C.teal})],{align:AlignmentType.CENTER})],{w:1600,fill:C.tealXL}),
            mkCell([bullet("Kesulitan mengikuti hitungan 1-8"),bullet("Gerakan kaki & tangan tidak sinkron"),bullet("Terlihat ragu-ragu & kurang percaya diri"),bullet("Belum pernah ikut senam/dance")],{w:3880}),
            mkCell(pRun("Stasiun 1 \u2014 Fokus gerak dasar & langkah sederhana dengan tempo lambat dan panduan video.",{size:20}),{w:3880,fill:C.tealXL}),
          ]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F33F} MENENGAH",{bold:true,size:20,color:C.orange})],{align:AlignmentType.CENTER})],{w:1600,fill:C.orangeXL}),
            mkCell([bullet("Dapat mengikuti langkah dasar"),bullet("Kesulitan pada kombinasi gerak tangan+kaki"),bullet("Kadang meleset dari irama saat perubahan gerak"),bullet("Pernah ikut senam/dance tingkat dasar")],{w:3880}),
            mkCell(pRun("Stasiun 2 \u2014 Kombinasi gerak berirama dengan alat sederhana & tempo sedang.",{size:20}),{w:3880,fill:C.orangeXL}),
          ]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F333} MAHIR",{bold:true,size:20,color:C.green})],{align:AlignmentType.CENTER})],{w:1600,fill:C.greenXL}),
            mkCell([bullet("Mengikuti irama dengan lancar & natural"),bullet("Menunjukkan ekspresi dan variasi gerak spontan"),bullet("Berinisiatif memimpin gerakan dalam kelompok"),bullet("Berpengalaman dalam tari/senam/dance")],{w:3880}),
            mkCell(pRun("Stasiun 3 \u2014 Kreasi rangkaian gerak berkelompok dengan pola lantai & variasi tempo.",{size:20}),{w:3880,fill:C.greenXL}),
          ]}),
        ]
      }),

      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 4: DIFERENSIASI
      // ════════════════════════════════════════════════════════════
      banner("4.  PEMBELAJARAN BERDIFERENSIASI",{fill:C.purple}),
      spacer(70),

      // 4A Konten
      subBanner("\u{1F4DA}  A. Diferensiasi KONTEN \u2014 Sumber Belajar Berjenjang",{fill:C.purpleXL,border:C.purple}),
      spacer(40),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1700,2700,2700,2260],
        rows:[
          new TableRow({children:[mkHdr("Level",{w:1700,fill:C.purple}),mkHdr("Media Utama",{w:2700,fill:C.purple}),mkHdr("Konten Spesifik",{w:2700,fill:C.purple}),mkHdr("Cara Akses",{w:2260,fill:C.purple})]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F331} PEMULA",{bold:true,size:20,color:C.teal})],{align:AlignmentType.CENTER,before:60,after:60})],{w:1700,fill:C.tealXL,vAlign:VerticalAlign.CENTER}),
            mkCell([p([run("\u25B6 Video Tutorial Gerak Dasar",{bold:true,size:20,color:C.teal})],{before:0,after:40}),pRun("Langkah kaki (march + step touch) & ayunan tangan dengan tempo lambat 60-80 BPM. Narasi hitungan 1-8 jelas.",{size:19})],{w:2700,fill:C.tealXL}),
            mkCell([bullet("Gerakan dipecah per hitungan"),bullet("Demo dari depan & samping"),bullet("Durasi video 3-5 menit"),bullet("Dapat diputar berulang")],{w:2700}),
            mkCell(pRun("TV/Proyektor kelas. Guru siapkan video offline.",{size:19}),{w:2260,fill:C.tealXL}),
          ]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F33F} MENENGAH",{bold:true,size:20,color:C.orange})],{align:AlignmentType.CENTER,before:60,after:60})],{w:1700,fill:C.orangeXL,vAlign:VerticalAlign.CENTER}),
            mkCell([p([run("\u{1F4CB} Kartu Infografis Gerak",{bold:true,size:20,color:C.orange})],{before:0,after:40}),pRun("Kombinasi gerak berirama: tangan+kaki+perpindahan arah. Disertai ikon posisi tubuh & arah gerak.",{size:19})],{w:2700,fill:C.orangeXL}),
            mkCell([bullet("Kartu A4 dicetak berwarna"),bullet("Diagram posisi tubuh tiap hitungan"),bullet("Kode warna tangan vs kaki"),bullet("Panduan perubahan arah")],{w:2700}),
            mkCell(pRun("Dicetak & dibagikan. Dijadikan panduan gerak mandiri.",{size:19}),{w:2260,fill:C.orangeXL}),
          ]}),
          new TableRow({children:[
            mkCell([p([run("\u{1F333} MAHIR",{bold:true,size:20,color:C.green})],{align:AlignmentType.CENTER,before:60,after:60})],{w:1700,fill:C.greenXL,vAlign:VerticalAlign.CENTER}),
            mkCell([p([run("\u{1F3AC} Artikel + Video Analisis",{bold:true,size:20,color:C.green})],{before:0,after:40}),pRun("Analisis pola lantai, variasi tempo, dan kreativitas gerak berirama modern (senam ritmik, aerobik, street dance).",{size:19})],{w:2700,fill:C.greenXL}),
            mkCell([bullet("Video pertunjukan profesional"),bullet("Artikel teknik pola lantai"),bullet("Lembar analisis kritis"),bullet("Referensi variasi tempo")],{w:2700}),
            mkCell(pRun("QR code ke sumber online. Peserta buat catatan kritis 5 poin.",{size:19}),{w:2260,fill:C.greenXL}),
          ]}),
        ]
      }),

      spacer(80),

      // 4B Proses
      subBanner("\u{1F504}  B. Diferensiasi PROSES \u2014 Sistem Stasiun Belajar (Rotasi 15 Menit)",{fill:C.blueXL,border:C.blue}),
      spacer(40),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3120,3120,3120],
        rows:[
          new TableRow({children:[mkHdr("STASIUN 1 \u2014 PEMULA",{w:3120,fill:C.teal}),mkHdr("STASIUN 2 \u2014 MENENGAH",{w:3120,fill:C.orange}),mkHdr("STASIUN 3 \u2014 MAHIR",{w:3120,fill:C.purple})]}),
          new TableRow({children:[
            mkCell([
              p([run("\u{1F3AF} FOKUS GERAK",{bold:true,size:20,color:C.teal})],{before:0,after:40}),
              pRun("Langkah dasar: march, step touch, heel touch. Ayunan tangan dua arah.",{size:19},{before:0,after:60}),
              p([run("\u{1F3B5} MUSIK",{bold:true,size:20,color:C.teal})],{before:0,after:40}),
              pRun("60-80 BPM. Hitungan 1-8 berulang. Guru memberi aba-aba lisan.",{size:19},{before:0,after:60}),
              p([run("\u2699 AKTIVITAS",{bold:true,size:20,color:C.teal})],{before:0,after:40}),
              bullet("Ikuti video gerak dasar step-by-step"),
              bullet("Latihan berpasangan: saling koreksi"),
              bullet("Guru mendampingi & memberi umpan balik"),
              p([run("\u{1F3C5} TARGET",{bold:true,size:20,color:C.teal})],{before:60,after:40}),
              pRun("Mampu lakukan 1x8 hitungan dengan benar.",{size:19}),
            ],{w:3120,fill:C.tealXL}),
            mkCell([
              p([run("\u{1F3AF} FOKUS GERAK",{bold:true,size:20,color:C.orange})],{before:0,after:40}),
              pRun("Kombinasi: langkah berpindah arah + ayunan tangan + alat (pita/syal).",{size:19},{before:0,after:60}),
              p([run("\u{1F3B5} MUSIK",{bold:true,size:20,color:C.orange})],{before:0,after:40}),
              pRun("80-100 BPM. Variasi 2x8 hitungan dengan perubahan arah.",{size:19},{before:0,after:60}),
              p([run("\u2699 AKTIVITAS",{bold:true,size:20,color:C.orange})],{before:0,after:40}),
              bullet("Latihan dengan kartu infografis"),
              bullet("Eksplorasi gerak dengan alat (pita/syal)"),
              bullet("Diskusi kelompok: evaluasi koordinasi"),
              p([run("\u{1F3C5} TARGET",{bold:true,size:20,color:C.orange})],{before:60,after:40}),
              pRun("Mampu lakukan 2x8 hitungan kombinasi.",{size:19}),
            ],{w:3120,fill:C.orangeXL}),
            mkCell([
              p([run("\u{1F3AF} FOKUS GERAK",{bold:true,size:20,color:C.purple})],{before:0,after:40}),
              pRun("Kreasi rangkaian gerak kelompok dengan pola lantai, variasi tempo, dan ekspresi.",{size:19},{before:0,after:60}),
              p([run("\u{1F3B5} MUSIK",{bold:true,size:20,color:C.purple})],{before:0,after:40}),
              pRun("100-120 BPM. Dinamis dengan perubahan birama dan tempo.",{size:19},{before:0,after:60}),
              p([run("\u2699 AKTIVITAS",{bold:true,size:20,color:C.purple})],{before:0,after:40}),
              bullet("Merancang pola lantai di kertas"),
              bullet("Kreasi 3-4 variasi gerakan unik kelompok"),
              bullet("Latihan penuh untuk demonstrasi akhir"),
              p([run("\u{1F3C5} TARGET",{bold:true,size:20,color:C.purple})],{before:60,after:40}),
              pRun("Tampilkan kreasi 4x8 hitungan dengan pola lantai.",{size:19}),
            ],{w:3120,fill:C.purpleXL}),
          ]}),
        ]
      }),

      spacer(80),

      // 4C Produk
      subBanner("\u{1F3AC}  C. Diferensiasi PRODUK \u2014 Demonstrasi Akhir",{fill:C.amberL,border:C.amber}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1600,4060,3700],
        rows:[
          new TableRow({children:[mkHdr("Level",{w:1600,fill:C.amber}),mkHdr("Bentuk Penampilan",{w:4060,fill:C.amber}),mkHdr("Kriteria Keberhasilan",{w:3700,fill:C.amber})]}),
          new TableRow({children:[
            mkCell(p([run("\u{1F331} PEMULA",{bold:true,size:20,color:C.teal})],{align:AlignmentType.CENTER}),{w:1600,fill:C.tealXL,vAlign:VerticalAlign.CENTER}),
            mkCell(pRun("Rangkaian gerak dasar 2x8 hitungan mengikuti musik. Boleh berpasangan atau bertiga.",{size:20}),{w:4060}),
            mkCell([bullet("Ketepatan langkah dasar"),bullet("Kesesuaian irama"),bullet("Keberanian tampil")],{w:3700,fill:C.tealXL}),
          ]}),
          new TableRow({children:[
            mkCell(p([run("\u{1F33F} MENENGAH",{bold:true,size:20,color:C.orange})],{align:AlignmentType.CENTER}),{w:1600,fill:C.orangeXL,vAlign:VerticalAlign.CENTER}),
            mkCell(pRun("Kombinasi gerak 3x8 hitungan dengan alat sederhana, perubahan arah minimal 2 kali.",{size:20}),{w:4060}),
            mkCell([bullet("Koordinasi tangan-kaki-alat"),bullet("Keluwesan gerak"),bullet("Perubahan arah tepat")],{w:3700,fill:C.orangeXL}),
          ]}),
          new TableRow({children:[
            mkCell(p([run("\u{1F333} MAHIR",{bold:true,size:20,color:C.green})],{align:AlignmentType.CENTER}),{w:1600,fill:C.greenXL,vAlign:VerticalAlign.CENTER}),
            mkCell(pRun("Kreasi gerak kelompok 4x8 hitungan dengan pola lantai sendiri, variasi tempo, dan ekspresi.",{size:20}),{w:4060}),
            mkCell([bullet("Kreativitas pola lantai"),bullet("Sinkronisasi kelompok"),bullet("Ekspresi & variasi tempo")],{w:3700,fill:C.greenXL}),
          ]}),
        ]
      }),

      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 5: LANGKAH KEGIATAN
      // ════════════════════════════════════════════════════════════
      banner("5.  LANGKAH KEGIATAN PEMBELAJARAN",{fill:C.teal}),
      spacer(70),

      // Pertemuan 1
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
        rows:[new TableRow({children:[new TableCell({
          children:[p([run("\u{1F4C5}  PERTEMUAN 1 \u2014 Eksplorasi & Pengenalan Gerak Berirama  (2 \u00D7 40 Menit)",{bold:true,size:24,color:C.white})],{before:90,after:90})],
          width:{size:9360,type:WidthType.DXA},shading:shade(C.teal),borders:bdrN(),margins:{top:60,bottom:60,left:180,right:180},
        })]})],
      }),
      spacer(70),

      subBanner("\u{1F3AC}  A. PENDAHULUAN (10 Menit)",{fill:C.tealXL,border:C.teal}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[680,1380,7300],
        rows:[
          new TableRow({children:[mkHdr("No",{w:680,fill:C.teal}),mkHdr("Waktu",{w:1380,fill:C.teal}),mkHdr("Deskripsi Kegiatan Guru & Peserta Didik",{w:7300,fill:C.teal})]}),
          ...[
            ["1","2 mnt","PEMBUKAAN: Guru mengucapkan salam, memimpin doa bersama, dan mengecek kehadiran. Peserta didik berbaris rapi di lapangan dengan jarak aman."],
            ["2","3 mnt","ICE BREAKING DANCE CHALLENGE: Guru memutar lagu populer dan memimpin 8 hitungan gerak sederhana yang fun. Peserta didik menirukan dengan bebas dan penuh semangat. Suasana dibangun ceria dan bebas penilaian."],
            ["3","2 mnt","APERSEPSI: Guru bertanya, 'Gerakan tadi terasa mudah atau sulit mengikuti iramanya? Kenapa?' Peserta berbagi pengalaman. Guru mengaitkan dengan kehidupan nyata (TikTok dance, senam pagi)."],
            ["4","2 mnt","PENYAMPAIAN TUJUAN: Guru menjelaskan tujuan pembelajaran hari ini secara singkat dan jelas. Menyampaikan alur kegiatan dan aturan stasiun belajar."],
            ["5","1 mnt","ASESMEN DIAGNOSTIK: Guru mengajukan Q1 dan Q2 secara lisan, mengamati respons, lalu meminta seluruh kelas mengikuti 8 hitungan gerak (Q3) sambil mencatat observasi."],
          ].map(([no,wkt,deskr],i)=>new TableRow({children:[
            mkCell(p([run(no,{bold:true,size:20})],{align:AlignmentType.CENTER}),{w:680,fill:i%2===0?C.tealXL:C.blueXL}),
            mkCell(p([run(wkt,{bold:true,size:20,color:C.teal})],{align:AlignmentType.CENTER}),{w:1380,fill:i%2===0?C.tealXL:C.blueXL}),
            mkCell(pRun(deskr,{size:20}),{w:7300,fill:i%2===0?C.white:C.grayXL}),
          ]})),
        ]
      }),
      spacer(70),

      subBanner("\u26A1  B. KEGIATAN INTI (25 Menit) \u2014 Diferensiasi Stasiun",{fill:C.orangeXL,border:C.orange}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[680,1380,7300],
        rows:[
          new TableRow({children:[mkHdr("No",{w:680,fill:C.orange}),mkHdr("Waktu",{w:1380,fill:C.orange}),mkHdr("Deskripsi Kegiatan",{w:7300,fill:C.orange})]}),
          ...[
            ["1","3 mnt","PENGELOMPOKAN: Guru membagi peserta didik ke 3 kelompok berdasarkan hasil diagnostik. Menjelaskan tugas tiap stasiun, target gerak, dan aturan rotasi dengan bahasa jelas dan singkat."],
            ["2","15 mnt","SESI STASIUN BELAJAR (3 stasiun berjalan serentak): Stasiun 1 (Pemula): Mengikuti video gerak dasar, latihan berpasangan, guru mendampingi. Stasiun 2 (Menengah): Eksplorasi kombinasi gerak dengan infografis & alat, diskusi kelompok. Stasiun 3 (Mahir): Merancang pola lantai & kreasi gerak kelompok, latihan penuh. Guru berpindah ke tiap stasiun setiap 5 menit untuk observasi & umpan balik."],
            ["3","4 mnt","JEDA REFLEKTIF: Guru menghentikan semua aktivitas. Musik dimatikan. Guru mengajukan 3 pertanyaan reflektif deep learning. Peserta diskusi berpasangan, lalu sharing ke kelas."],
            ["4","3 mnt","MINI PENAMPILAN: Satu perwakilan tiap stasiun menampilkan 1x8 hitungan gerak. Seluruh kelas memberikan apresiasi tepuk tangan. Guru memberi umpan balik positif & konstruktif."],
          ].map(([no,wkt,deskr],i)=>new TableRow({children:[
            mkCell(p([run(no,{bold:true,size:20})],{align:AlignmentType.CENTER}),{w:680,fill:i%2===0?C.orangeXL:C.amberL}),
            mkCell(p([run(wkt,{bold:true,size:20,color:C.orange})],{align:AlignmentType.CENTER}),{w:1380,fill:i%2===0?C.orangeXL:C.amberL}),
            mkCell(pRun(deskr,{size:20}),{w:7300,fill:i%2===0?C.white:C.grayXL}),
          ]})),
        ]
      }),
      spacer(60),

      callout(
        p([run("\u{1F4AC}  Pertanyaan Reflektif Deep Learning \u2014 Pertemuan 1",{bold:true,size:21,color:C.white})],{before:0,after:0}),
        [
          p([run("1. ",{bold:true,size:21,color:C.blue}),run('"Mengapa gerakan HARUS mengikuti tempo musik? Apa yang terjadi jika tidak mengikutinya?"',{italic:true,size:21})],{before:0,after:40}),
          p([run("2. ",{bold:true,size:21,color:C.blue}),run('"Apa yang terjadi pada kelompok jika satu orang tidak sinkron? Bagaimana solusinya?"',{italic:true,size:21})],{before:0,after:40}),
          p([run("3. ",{bold:true,size:21,color:C.blue}),run('"Bagaimana perasaan tubuhmu setelah bergerak mengikuti irama? Bagian mana yang paling bekerja keras?"',{italic:true,size:21})],{before:0,after:0}),
        ],
        {fill:C.blueXL,accent:C.blue}
      ),
      spacer(70),

      subBanner("\u{1F319}  C. PENUTUP (5 Menit)",{fill:C.purpleXL,border:C.purple}),
      bullet("PENDINGINAN: Peregangan ringan seluruh tubuh diiringi musik instrumental lembut selama 3 menit. Guru mengajak peserta merasakan relaksasi otot secara sadar (mindful stretching)."),
      bullet("REFLEKSI EXIT TICKET: Peserta menyampaikan satu hal berkesan: 'Saya belajar bahwa...' Guru mencatat poin menarik dari respons peserta."),
      bullet("PENGUATAN & MOTIVASI: Guru mengapresiasi seluruh kelas atas usaha dan semangat. Menyampaikan bahwa pertemuan berikutnya ada sesi demonstrasi dan kreasi yang lebih seru."),

      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // Pertemuan 2
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
        rows:[new TableRow({children:[new TableCell({
          children:[p([run("\u{1F4C5}  PERTEMUAN 2 \u2014 Pendalaman, Kreasi & Demonstrasi Akhir  (2 \u00D7 40 Menit)",{bold:true,size:24,color:C.white})],{before:90,after:90})],
          width:{size:9360,type:WidthType.DXA},shading:shade(C.orange),borders:bdrN(),margins:{top:60,bottom:60,left:180,right:180},
        })]})],
      }),
      spacer(70),

      subBanner("\u{1F3AC}  A. PENDAHULUAN (8 Menit)",{fill:C.orangeXL,border:C.orange}),
      bullet("REVIEW: Guru bertanya, 'Siapa yang masih ingat 3 gerakan dasar dari pertemuan lalu?' Peserta berbagi jawaban secara cepat."),
      bullet("ICE BREAKING ROTATE: Satu peserta memimpin 4 hitungan gerak favoritnya, bergantian 3 peserta. Kelas mengikuti dengan semangat."),
      bullet("MOTIVASI: Guru menyampaikan bahwa hari ini adalah sesi puncak: kreasi dan demonstrasi. 'Kalian sudah berlatih keras, sekarang saatnya bersinar!'"),
      bullet("TEKNIS: Guru menjelaskan aturan demonstrasi akhir, rubrik penilaian, dan alokasi waktu penampilan tiap kelompok."),
      spacer(60),

      subBanner("\u26A1  B. KEGIATAN INTI (27 Menit)",{fill:C.tealXL,border:C.teal}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[680,1380,7300],
        rows:[
          new TableRow({children:[mkHdr("No",{w:680,fill:C.teal}),mkHdr("Waktu",{w:1380,fill:C.teal}),mkHdr("Deskripsi Kegiatan",{w:7300,fill:C.teal})]}),
          ...[
            ["1","10 mnt","STASIUN LANJUTAN: Peserta kembali ke stasiun masing-masing untuk sesi latihan terakhir. Fokus pada aspek yang perlu diperbaiki berdasarkan refleksi pertemuan 1. Guru memberi umpan balik individual yang spesifik: tempo, koordinasi, dan ekspresi."],
            ["2","7 mnt","LATIHAN KOLABORATIF: Kelompok Pemula + Menengah saling menonton dan memberi umpan balik konstruktif: 'Satu hal yang kamu lakukan dengan baik adalah... Satu saran perbaikan adalah...' Kelompok Mahir menyempurnakan pola lantai."],
            ["3","10 mnt","SESI DEMONSTRASI AKHIR: Tiap kelompok tampil sesuai level. Penonton mengisi lembar peer assessment sederhana dan memberi apresiasi. Guru mengobservasi dan mencatat skor pada rubrik penilaian."],
          ].map(([no,wkt,deskr],i)=>new TableRow({children:[
            mkCell(p([run(no,{bold:true,size:20})],{align:AlignmentType.CENTER}),{w:680,fill:i%2===0?C.tealXL:C.blueXL}),
            mkCell(p([run(wkt,{bold:true,size:20,color:C.teal})],{align:AlignmentType.CENTER}),{w:1380,fill:i%2===0?C.tealXL:C.blueXL}),
            mkCell(pRun(deskr,{size:20}),{w:7300,fill:i%2===0?C.white:C.grayXL}),
          ]})),
        ]
      }),
      spacer(60),

      subBanner("\u{1F319}  C. PENUTUP (5 Menit)",{fill:C.purpleXL,border:C.purple}),
      bullet("PENDINGINAN MINDFUL: Peregangan seluruh tubuh dengan musik instrumental. Guru mengajak peserta merasakan setiap regangan secara sadar."),
      bullet("REFLEKSI AKHIR: Peserta mengisi Bagian E LKPD (Evaluasi Diri). 3-4 peserta berbagi refleksi singkat kepada kelas."),
      bullet("PENGUATAN KONSEP: Guru merangkum poin kunci: koordinasi, tempo, kerja sama, dan ekspresi diri melalui gerak."),
      bullet("PENUTUP: Guru mengapresiasi partisipasi aktif seluruh peserta. Doa penutup bersama."),

      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 6: LKPD
      // ════════════════════════════════════════════════════════════
      banner("6.  LEMBAR KERJA PESERTA DIDIK (LKPD)",{fill:C.purple}),
      spacer(40),

      // LKPD Header
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
        rows:[new TableRow({children:[new TableCell({
          children:[
            p([run("\u{1F3B6} \u{1F483}  LEMBAR KERJA PESERTA DIDIK  \u{1F57A} \u{1F3B6}",{bold:true,size:30,color:C.white})],{align:AlignmentType.CENTER,before:70,after:30}),
            p([run("AKTIVITAS GERAK BERIRAMA  |  PJOK Kelas 8  |  Fase D",{size:21,color:C.purpleL})],{align:AlignmentType.CENTER,before:0,after:30}),
            p([run('"Setiap Gerakan adalah Ekspresi Dirimu!"',{italic:true,size:19,color:C.amberL})],{align:AlignmentType.CENTER,before:0,after:70}),
          ],
          width:{size:9360,type:WidthType.DXA},shading:shade(C.purple),borders:bdrN(),margins:{top:50,bottom:50,left:180,right:180},
        })]})],
      }),

      // Identity fields
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1400,2860,1400,3700],
        rows:[
          new TableRow({children:[
            mkCell(pRun("Nama:",{bold:true,size:20,color:C.purple}),{w:1400,fill:C.purpleXL}),
            mkCell(pRun("________________________________",{size:20,color:C.grayL}),{w:2860}),
            mkCell(pRun("Kelas:",{bold:true,size:20,color:C.purple}),{w:1400,fill:C.purpleXL}),
            mkCell(pRun("_______________",{size:20,color:C.grayL}),{w:3700}),
          ]}),
          new TableRow({children:[
            mkCell(pRun("Tanggal:",{bold:true,size:20,color:C.purple}),{w:1400,fill:C.purpleXL}),
            mkCell(pRun("________________________________",{size:20,color:C.grayL}),{w:2860}),
            mkCell(pRun("Kelompok:",{bold:true,size:20,color:C.purple}),{w:1400,fill:C.purpleXL}),
            mkCell(pRun("_______________",{size:20,color:C.grayL}),{w:3700}),
          ]}),
        ]
      }),
      spacer(70),

      // LKPD A
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({children:[p([run("\u{1F441}  BAGIAN A \u2014 PENGAMATAN GERAK BERIRAMA",{bold:true,size:21,color:C.white})],{before:70,after:70})],width:{size:9360,type:WidthType.DXA},shading:shade(C.teal),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160}})]})]},
      spacer(40),
      pRun("Amati demonstrasi gerak dari guru atau video, lalu isi tabel berikut dengan teliti:"),
      spacer(50),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[560,1900,1800,1700,1400,1700 +300],
        rows:[
          new TableRow({children:[mkHdr("No",{w:560,fill:C.teal}),mkHdr("Nama Gerakan",{w:1900,fill:C.teal}),mkHdr("Bagian Tubuh",{w:1800,fill:C.teal}),mkHdr("Tempo Musik",{w:1700,fill:C.teal}),mkHdr("Kesesuaian Irama",{w:1400,fill:C.teal}),mkHdr("Catatan Observasiku",{w:2000,fill:C.teal})]}),
          ...[1,2,3,4,5].map((n,i)=>new TableRow({children:[
            mkCell(p([run(`${n}`,{bold:true,size:20})],{align:AlignmentType.CENTER}),{w:560,fill:i%2===0?C.tealXL:C.white}),
            mkCell(pRun(".......................",{size:20,color:C.grayM}),{w:1900,fill:i%2===0?C.tealXL:C.white}),
            mkCell(pRun(".......................",{size:20,color:C.grayM}),{w:1800,fill:i%2===0?C.tealXL:C.white}),
            mkCell(pRun(".......................",{size:20,color:C.grayM}),{w:1700,fill:i%2===0?C.tealXL:C.white}),
            mkCell([p([run("\u2610 Sangat Sesuai",{size:17,color:C.green})],{before:0,after:20}),p([run("\u2610 Cukup Sesuai",{size:17,color:C.orange})],{before:0,after:20}),p([run("\u2610 Belum Sesuai",{size:17,color:C.red})],{before:0,after:0})],{w:1400,fill:i%2===0?C.tealXL:C.white}),
            mkCell(pRun(".......................",{size:20,color:C.grayM}),{w:2000,fill:i%2===0?C.tealXL:C.white}),
          ]})),
        ]
      }),
      spacer(50),
      p([run("\u{1F50D} Dari 5 gerakan yang kamu amati, gerakan mana yang PALING SULIT? Mengapa?",{bold:true,size:20,color:C.teal})],{before:0,after:40}),
      pRun("Jawab: ______________________________________________________________________________________",{size:20,color:C.grayL}),
      pRun("       ______________________________________________________________________________________",{size:20,color:C.grayL}),
      spacer(70),

      // LKPD B
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({children:[p([run("\u2705  BAGIAN B \u2014 CHECKLIST KOORDINASI GERAK",{bold:true,size:21,color:C.white})],{before:70,after:70})],width:{size:9360,type:WidthType.DXA},shading:shade(C.orange),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160}})]})]},
      spacer(40),
      pRun("Beri tanda (\u2713) pada kolom yang PALING SESUAI dengan kemampuanmu saat ini. Jujur ya!"),
      spacer(50),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[4200,1720,1720,1720],
        rows:[
          new TableRow({children:[mkHdr("Aspek Koordinasi Gerak",{w:4200,fill:C.orange}),mkHdr("\u2B50\u2B50\u2B50 Sudah Mahir",{w:1720,fill:C.green}),mkHdr("\u2B50\u2B50 Sedang Berlatih",{w:1720,fill:C.orange}),mkHdr("\u2B50 Perlu Bantuan",{w:1720,fill:C.red})]}),
          ...[
            "Langkah kaki mengikuti irama (march, step touch)",
            "Ayunan tangan selaras dengan langkah kaki",
            "Keseimbangan tubuh saat berpindah arah",
            "Ketepatan gerak sesuai hitungan 1-8",
            "Gerak badan terlihat luwes dan rileks (tidak kaku)",
            "Sinkronisasi dengan gerakan teman sekelompok",
            "Perpindahan formasi / pola lantai yang teratur",
            "Ekspresi wajah positif dan percaya diri saat tampil",
          ].map((item,i)=>new TableRow({children:[
            mkCell(pRun(item,{size:20}),{w:4200,fill:i%2===0?C.orangeXL:C.white}),
            mkCell(p([run("\u2610",{size:24,color:C.green})],{align:AlignmentType.CENTER}),{w:1720,fill:C.greenXL}),
            mkCell(p([run("\u2610",{size:24,color:C.orange})],{align:AlignmentType.CENTER}),{w:1720,fill:C.orangeXL}),
            mkCell(p([run("\u2610",{size:24,color:C.red})],{align:AlignmentType.CENTER}),{w:1720,fill:C.redXL}),
          ]})),
        ]
      }),
      spacer(40),
      p([run("\u{1F4A1} Tuliskan 2 aspek yang paling kamu banggakan dan 1 aspek yang ingin kamu tingkatkan:",{bold:true,size:20,color:C.orange})],{before:0,after:40}),
      pRun("Bangga: (1) __________________ (2) __________________     Ingin tingkatkan: __________________",{size:20,color:C.grayM}),
      spacer(70),

      new Paragraph({children:[new PageBreak()]}),

      // LKPD C
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({children:[p([run("\u{1F497}  BAGIAN C \u2014 REFLEKSI PERASAAN SETELAH BERGERAK",{bold:true,size:21,color:C.white})],{before:70,after:70})],width:{size:9360,type:WidthType.DXA},shading:shade(C.purple),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160}})]})]},
      spacer(40),
      pRun("Pilih SATU wajah yang paling menggambarkan perasaanmu hari ini dengan melingkari pilihanmu:"),
      spacer(50),

      // Mood selector
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1872,1872,1872,1872,1872],
        rows:[new TableRow({children:[
          new TableCell({children:[p([run("\u{1F604}",{font:"Segoe UI Emoji",size:48})],{align:AlignmentType.CENTER,before:40,after:20}),p([run("SANGAT SENANG",{bold:true,size:17,color:C.green})],{align:AlignmentType.CENTER,before:0,after:40})],width:{size:1872,type:WidthType.DXA},shading:shade(C.greenXL),borders:bdrT(C.green,6),margins:{top:40,bottom:40,left:80,right:80}}),
          new TableCell({children:[p([run("\u{1F60A}",{font:"Segoe UI Emoji",size:48})],{align:AlignmentType.CENTER,before:40,after:20}),p([run("SENANG",{bold:true,size:17,color:C.teal})],{align:AlignmentType.CENTER,before:0,after:40})],width:{size:1872,type:WidthType.DXA},shading:shade(C.tealXL),borders:bdrT(C.teal,6),margins:{top:40,bottom:40,left:80,right:80}}),
          new TableCell({children:[p([run("\u{1F610}",{font:"Segoe UI Emoji",size:48})],{align:AlignmentType.CENTER,before:40,after:20}),p([run("BIASA SAJA",{bold:true,size:17,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:40})],width:{size:1872,type:WidthType.DXA},shading:shade(C.grayXL),borders:bdrT(C.grayM,6),margins:{top:40,bottom:40,left:80,right:80}}),
          new TableCell({children:[p([run("\u{1F615}",{font:"Segoe UI Emoji",size:48})],{align:AlignmentType.CENTER,before:40,after:20}),p([run("KURANG PD",{bold:true,size:17,color:C.orange})],{align:AlignmentType.CENTER,before:0,after:40})],width:{size:1872,type:WidthType.DXA},shading:shade(C.orangeXL),borders:bdrT(C.orange,6),margins:{top:40,bottom:40,left:80,right:80}}),
          new TableCell({children:[p([run("\u{1F62C}",{font:"Segoe UI Emoji",size:48})],{align:AlignmentType.CENTER,before:40,after:20}),p([run("SANGAT SULIT",{bold:true,size:17,color:C.red})],{align:AlignmentType.CENTER,before:0,after:40})],width:{size:1872,type:WidthType.DXA},shading:shade(C.redXL),borders:bdrT(C.red,6),margins:{top:40,bottom:40,left:80,right:80}}),
        ]})]
      }),
      spacer(60),
      ...[
        ["\u{1F3AF}","Apa momen paling berkesan yang kamu alami dalam pembelajaran hari ini? Ceritakan!"],
        ["\u{1F914}","Gerakan apa yang paling menantang? Strategi apa yang kamu gunakan untuk mengatasinya?"],
        ["\u{1F4AA}","Apa yang akan kamu lakukan berbeda jika bisa mengulang pelajaran ini dari awal?"],
        ["\u{1F91D}","Bagaimana perasaanmu saat bergerak bersama kelompok? Adakah momen kerja sama yang berkesan?"],
        ["\u{1F9D8}","Setelah bergerak mengikuti irama, bagian tubuh mana yang paling kamu rasakan aktif bekerja?"],
      ].flatMap(([icon,q])=>[
        p([run(`${icon} ${q}`,{bold:true,size:20,color:C.purple})],{before:60,after:40}),
        pRun("Jawab: ______________________________________________________________________________",{size:20,color:C.grayL},{before:0,after:20}),
        pRun("       ______________________________________________________________________________",{size:20,color:C.grayL},{before:0,after:0}),
      ]),
      spacer(70),

      // LKPD D
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({children:[p([run("\u{1F3A8}  BAGIAN D \u2014 TANTANGAN KREATIVITAS GERAK",{bold:true,size:21,color:C.white})],{before:70,after:70})],width:{size:9360,type:WidthType.DXA},shading:shade(C.green),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160}})]})]},
      spacer(40),

      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({
        children:[
          p([run("\u{1F525}  TANTANGAN DANCE CHALLENGE KELOMPOK  \u{1F525}",{bold:true,size:24,color:C.amber})],{align:AlignmentType.CENTER,before:60,after:40}),
          p([run("Bersama kelompokmu, ciptakan MINI DANCE CHALLENGE 2x8 hitungan yang UNIK & KREATIF!",{size:21,color:C.black})],{align:AlignmentType.CENTER,before:0,after:50}),
          ...[
            ["1","Gunakan MINIMAL 3 jenis gerakan berbeda (tangan, kaki, dan badan)."],
            ["2","SEMUA anggota bergerak serempak dan kompak."],
            ["3","Pilih musikmu sendiri, sesuaikan gerakan dengan temponya."],
            ["4","Tambahkan MINIMAL 1 gerakan kreatif ciptaan sendiri yang BELUM pernah ada!"],
            ["5","Beri nama dance challengemu yang keren dan mudah diingat."],
          ].map(([n,t])=>p([run(`${n}. `,{bold:true,size:21,color:C.amber}),run(t,{size:21,color:C.black})],{before:20,after:20})),
          spacer(30),
        ],
        width:{size:9360,type:WidthType.DXA},shading:shade(C.amberL),borders:bdrT(C.amber,10),margins:{top:50,bottom:50,left:180,right:180},
      })]})]}),

      spacer(50),
      p([run("Nama Dance Challenge Kelompokku: ",{bold:true,size:21,color:C.green}),run('"_________________________________"',{size:21,color:C.grayM})],{before:0,after:50}),
      pRun("Rancangan Gerakan (deskripsi tiap hitungan atau sketsa gerakanmu!):",{bold:true,size:20,color:C.green},{before:0,after:40}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[900,3480,480,900,3600],
        rows:[
          new TableRow({children:[mkHdr("Hitungan",{w:900,fill:C.green}),mkHdr("Deskripsi / Sketsa Gerakan",{w:3480,fill:C.green}),mkHdr("",{w:480,fill:C.white}),mkHdr("Hitungan",{w:900,fill:C.green}),mkHdr("Deskripsi / Sketsa Gerakan",{w:3600,fill:C.green})]}),
          ...[1,2,3,4].map(n=>new TableRow({children:[
            mkCell(p([run(`${n}`,{bold:true,size:21,color:C.green})],{align:AlignmentType.CENTER}),{w:900,fill:C.greenXL}),
            mkCell(new Paragraph({children:[run("  ",{size:44})],spacing:{before:80,after:80}}),{w:3480}),
            mkCell(new Paragraph({children:[run("")]}),{w:480,borders:bdrN()}),
            mkCell(p([run(`${n+4}`,{bold:true,size:21,color:C.green})],{align:AlignmentType.CENTER}),{w:900,fill:C.greenXL}),
            mkCell(new Paragraph({children:[run("  ",{size:44})],spacing:{before:80,after:80}}),{w:3600}),
          ]})),
        ]
      }),
      spacer(50),
      p([run("\u{1F5FA}  Rancangan Pola Lantai (gambar dan beri keterangan formasi kelompokmu!):",{bold:true,size:20,color:C.green})],{before:0,after:40}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[4580,4780],rows:[new TableRow({children:[
        mkCell([new Paragraph({children:[run("  ",{size:180})],spacing:{before:0,after:0}}),p([run("Area Gambar Pola Lantai",{size:17,color:C.grayL,italic:true})],{align:AlignmentType.CENTER,before:0,after:0}),new Paragraph({children:[run("  ",{size:180})],spacing:{before:0,after:0}})],{w:4580,borders:bdrT(C.green,6)}),
        mkCell([pRun("Keterangan Formasi:",{bold:true,size:20,color:C.green},{before:0,after:50}),
          ...["A = ________________","B = ________________","C = ________________","Arah gerak: ________________"].map(t=>pRun(t,{size:19,color:C.grayM},{before:20,after:20})),
          spacer(50),pRun("Jumlah hitungan: ______ x 8",{size:19,color:C.grayM},{before:0,after:20}),pRun("Tempo musik: __________ BPM",{size:19,color:C.grayM},{before:0,after:0})],{w:4780}),
      ]})]},
      spacer(70),

      new Paragraph({children:[new PageBreak()]}),

      // LKPD E
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({children:[p([run("\u2B50  BAGIAN E \u2014 EVALUASI DIRI & RENCANA PERBAIKAN",{bold:true,size:21,color:C.white})],{before:70,after:70})],width:{size:9360,type:WidthType.DXA},shading:shade(C.gray),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160}})]})]},
      spacer(40),
      pRun("Lingkari jumlah bintang yang sesuai (1=perlu banyak latihan, 5=sudah sangat bisa), lalu tulis rencanamu!"),
      spacer(50),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3200,2560,3600],
        rows:[
          new TableRow({children:[mkHdr("Aspek yang Dinilai",{w:3200,fill:C.gray}),mkHdr("Nilai Diri (lingkari)",{w:2560,fill:C.gray}),mkHdr("Rencana Aksi Perbaikanku",{w:3600,fill:C.gray})]}),
          ...[
            "Kemampuan mengikuti tempo musik",
            "Koordinasi gerak tangan & kaki",
            "Keluwesan dan kenyamanan bergerak",
            "Kerja sama & sinkronisasi kelompok",
            "Kepercayaan diri saat tampil",
            "Kreativitas dalam variasi gerak",
            "Disiplin mengikuti aturan pembelajaran",
          ].map((item,i)=>new TableRow({children:[
            mkCell(pRun(item,{size:20}),{w:3200,fill:i%2===0?C.grayXL:C.white}),
            mkCell(p([run("1  \u2B50  2  \u2B50  3  \u2B50  4  \u2B50  5",{font:"Segoe UI Emoji",size:17,color:C.amber})],{align:AlignmentType.CENTER}),{w:2560,fill:C.yellowL}),
            mkCell(pRun("Saya akan ___________________",{size:19,color:C.grayM}),{w:3600,fill:i%2===0?C.grayXL:C.white}),
          ]})),
        ]
      }),
      spacer(60),

      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({
        children:[
          p([run("\u{1F48C}  SURAT UNTUKKU SENDIRI",{bold:true,size:21,color:C.purple})],{before:0,after:40}),
          pRun("Tuliskan satu kalimat penyemangat untuk dirimu sendiri tentang perjalanan belajar gerak beriramamu:",{size:20,color:C.black},{before:0,after:40}),
          pRun('"Aku bangga karena _______________________________________________________________"',{size:21,italic:true,color:C.grayM},{before:0,after:40}),
          pRun('"Aku berjanji akan terus __________________________________________________________"',{size:21,italic:true,color:C.grayM},{before:0,after:0}),
        ],
        width:{size:9360,type:WidthType.DXA},shading:shade(C.purpleXL),borders:bdrT(C.purple,8),margins:{top:100,bottom:100,left:180,right:180},
      })]})]}),
      spacer(60),

      // Closing
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],rows:[new TableRow({children:[new TableCell({
        children:[
          p([run("\u{1F57A}  Keep Moving, Keep Smiling, Keep Growing!  \u{1F483}",{bold:true,size:26,color:C.white})],{align:AlignmentType.CENTER,before:70,after:30}),
          p([run("Setiap langkah yang kamu ambil, setiap hitungan yang kamu ikuti \u2014 itu semua adalah kemajuanmu!",{size:19,color:C.tealXL,italic:true})],{align:AlignmentType.CENTER,before:0,after:70}),
        ],
        width:{size:9360,type:WidthType.DXA},shading:shade(C.tealD),borders:bdrN(),margins:{top:50,bottom:50,left:160,right:160},
      })]})]}),
      spacer(60),
      new Paragraph({children:[new PageBreak()]}),

      // ════════════════════════════════════════════════════════════
      // BAGIAN 7: RUBRIK ASESMEN
      // ════════════════════════════════════════════════════════════
      banner("7.  RUBRIK ASESMEN SUMATIF",{fill:C.tealD}),
      spacer(70),

      // 7A Rubrik Demonstrasi
      subBanner("\u{1F4CA}  A. Rubrik Penilaian Demonstrasi Akhir \u2014 Skala 1-4",{fill:C.tealXL,border:C.teal}),
      spacer(40),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2100,1815,1815,1815,1815],
        rows:[
          new TableRow({children:[
            mkHdr("Aspek Penilaian",{w:2100,fill:C.tealD}),
            mkHdr("4 \u2014 Sangat Baik",{w:1815,fill:C.green}),
            mkHdr("3 \u2014 Baik",{w:1815,fill:C.teal}),
            mkHdr("2 \u2014 Cukup",{w:1815,fill:C.orange}),
            mkHdr("1 \u2014 Perlu Bimbingan",{w:1815,fill:C.red}),
          ]}),
          ...[
            ["1. Ketepatan Tempo","Selalu mengikuti tempo musik dengan sangat tepat sepanjang penampilan.","Sebagian besar (>75%) gerakan tepat sesuai tempo.","Tepat tempo hanya pada bagian tertentu (<50%).","Belum mampu menyesuaikan gerak dengan tempo."],
            ["2. Koordinasi Gerak","Tangan, kaki, dan tubuh terkoordinasi sempurna tanpa kesalahan berarti.","Koordinasi baik dengan 1-2 kesalahan minor.","Koordinasi belum konsisten; sering tidak sesuai.","Koordinasi sangat kurang; tangan & kaki tidak sinkron."],
            ["3. Keluwesan Gerak","Gerakan sangat luwes, natural, dan penuh ekspresi.","Gerakan cukup luwes dengan sedikit kekakuan ringan.","Gerakan masih terlihat kaku dan kurang rileks.","Gerakan sangat kaku dan terlihat tegang."],
            ["4. Kerja Sama & Sinkronisasi","Seluruh anggota kelompok bergerak serentak dan sinkron sempurna.","Sebagian besar (>75%) gerakan sinkron.","Sering terjadi ketidaksinkronan antar anggota.","Tidak ada sinkronisasi; tiap anggota bergerak sendiri."],
            ["5. Ekspresi & Kepercayaan Diri","Ekspresi sangat positif, penuh energi, dan sangat percaya diri.","Percaya diri dengan ekspresi yang baik dan konsisten.","Ekspresi kurang; masih terlihat ragu-ragu.","Sangat tidak percaya diri; ekspresi tegang."],
            ["6. Kreativitas (Mahir)","Variasi gerak sangat kreatif, pola lantai menarik & inovatif.","Ada kreativitas cukup jelas dalam variasi & formasi.","Kreativitas terbatas; variasi gerakan minimal.","Tidak ada variasi kreatif; hanya meniru gerak dasar."],
          ].map(([aspek,...cols],i)=>new TableRow({children:[
            mkCell(pRun(aspek,{bold:true,size:19,color:C.teal}),{w:2100,fill:i%2===0?C.tealXL:C.blueXL}),
            ...cols.map((c,ci)=>mkCell(pRun(c,{size:18}),{w:1815,fill:i%2===0?[C.greenXL,C.tealXL,C.orangeXL,C.redXL][ci]:[C.white,C.white,C.white,C.white][ci]})),
          ]})),
          new TableRow({children:[
            mkHdr("SKOR MAKSIMAL",{w:2100,fill:C.teal}),
            new TableCell({
              children:[p([run("Nilai Akhir = (Total Skor / 24) \u00D7 100",{bold:true,size:20,color:C.teal})],{before:60,after:40}),p([run("Catatan: Aspek ke-6 HANYA untuk kelompok Stasiun Mahir.",{size:17,color:C.grayM,italic:true})],{before:0,after:60})],
              columnSpan:4,width:{size:7260,type:WidthType.DXA},
              shading:shade(C.tealXL),borders:bdrT(C.teal,6),margins:{top:70,bottom:70,left:140,right:140},
            }),
          ]}),
        ]
      }),
      spacer(70),

      // 7B Lembar Observasi
      subBanner("\u{1F4CB}  B. Lembar Observasi Guru \u2014 Asesmen Proses",{fill:C.blueXL,border:C.blue}),
      spacer(40),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1400,1000,4060,2900],
        rows:[
          new TableRow({children:[mkHdr("No. Absen",{w:1400,fill:C.blue}),mkHdr("Level",{w:1000,fill:C.blue}),mkHdr("Catatan Observasi Proses Belajar",{w:4060,fill:C.blue}),mkHdr("Umpan Balik Diberikan",{w:2900,fill:C.blue})]}),
          ...[1,2,3,4,5,6,7,8].map((_,i)=>new TableRow({children:[
            mkCell(p([run(`${i+1}`,{size:20})],{align:AlignmentType.CENTER}),{w:1400,fill:i%2===0?C.blueXL:C.white}),
            mkCell(p([run("P/M/A",{size:18,color:C.grayM,italic:true})],{align:AlignmentType.CENTER}),{w:1000,fill:i%2===0?C.blueXL:C.white}),
            mkCell(pRun("...............................................................",{size:19,color:C.grayL}),{w:4060,fill:i%2===0?C.blueXL:C.white}),
            mkCell(pRun("...............................................",{size:19,color:C.grayL}),{w:2900,fill:i%2===0?C.blueXL:C.white}),
          ]})),
        ]
      }),
      spacer(30),
      pRun("P = Pemula    M = Menengah    A = Mahir (Advanced)",{size:17,color:C.grayM,italic:true},{before:0,after:0}),
      spacer(70),

      // 7C Rekap Nilai
      subBanner("\u{1F3AF}  C. Rekap Nilai & Catatan Akhir",{fill:C.amberL,border:C.amber}),
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3600,1800,3960],
        rows:[
          new TableRow({children:[mkHdr("Nama Peserta Didik",{w:3600,fill:C.amber}),mkHdr("Nilai Akhir (/100)",{w:1800,fill:C.amber}),mkHdr("Catatan & Rekomendasi Guru",{w:3960,fill:C.amber})]}),
          ...[1,2,3,4,5,6].map((_,i)=>new TableRow({children:[
            mkCell(pRun("___________________________",{size:20,color:C.grayL}),{w:3600,fill:i%2===0?C.amberL:C.white}),
            mkCell(p([run("______",{size:20,color:C.grayL})],{align:AlignmentType.CENTER}),{w:1800,fill:i%2===0?C.amberL:C.white}),
            mkCell(pRun("____________________________",{size:20,color:C.grayL}),{w:3960,fill:i%2===0?C.amberL:C.white}),
          ]})),
        ]
      }),

      spacer(100),

      // Tanda tangan
      new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[4680,4680],
        rows:[new TableRow({children:[
          mkCell([
            p([run("Mengetahui,",{size:20,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Kepala Sekolah",{bold:true,size:21,color:C.gray})],{align:AlignmentType.CENTER,before:0,after:200}),
            p([run("( ______________________________ )",{size:20,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("NIP. ___________________________",{size:19,color:C.grayL})],{align:AlignmentType.CENTER,before:0,after:0}),
          ],{w:4680,borders:bdrN()}),
          mkCell([
            p([run("_______________, ___ ________ 20___",{size:20,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("Guru Mata Pelajaran PJOK",{bold:true,size:21,color:C.gray})],{align:AlignmentType.CENTER,before:0,after:200}),
            p([run("( ______________________________ )",{size:20,color:C.grayM})],{align:AlignmentType.CENTER,before:0,after:20}),
            p([run("NIP. ___________________________",{size:19,color:C.grayL})],{align:AlignmentType.CENTER,before:0,after:0}),
          ],{w:4680,borders:bdrN()}),
        ]})]
      }),

    ],
  }],
});

Packer.toBuffer(doc).then(buffer=>{
  fs.writeFileSync("/mnt/user-data/outputs/Modul_Ajar_PJOK_K8_DeepLearning_V2.docx",buffer);
  console.log("SUCCESS!");
}).catch(err=>{
  console.error("ERROR:",err.message);
  process.exit(1);
});
"""

# Unicode-escape all non-ASCII characters (emoji etc.)
import re

def escape_unicode(s):
    result = []
    i = 0
    while i < len(s):
        c = s[i]
        cp = ord(c)
        if cp > 127:
            result.append(f'\\u{{{cp:X}}}')
        else:
            result.append(c)
        i += 1
    return ''.join(result)

escaped = escape_unicode(JS)

with open('modul_v2_clean.js', 'w', encoding='ascii') as f:
    f.write(escaped)

print(f"Written {len(escaped)} chars to modul_v2_clean.js")
