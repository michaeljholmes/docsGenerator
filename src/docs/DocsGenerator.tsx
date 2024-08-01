import { saveAs } from "file-saver";
import { AppBar, Button, Checkbox, FormControl, FormControlLabel, FormGroup, InputLabel, ListItemText, MenuItem, OutlinedInput, Select, SelectChangeEvent, Stack, Toolbar, Typography } from "@mui/material";
import {
    AlignmentType,
    Document,
    Footer,
    FrameAnchorType,
    Header,
    HeadingLevel,
    HorizontalPositionAlign,
    IRunOptions,
    ImageRun,
    LineRuleType,
    Packer,
    PageBreak,
    PageNumber,
    Paragraph,
    SectionType,
    ShadingType,
    TextRun,
    VerticalPositionAlign,
  } from "docx";
import { useState } from "react";
import { rem } from "polished";
import { Field, Form, Formik } from "formik";
import { object, string } from "yup";
import { TextField } from 'formik-mui';

const imageBase64Data = 'data:image/png;base64,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';

type BlockType = "Normal";

// TODO - add in required
type CustomParagraph = Required<Pick<IRunOptions, 'text'>> & IRunOptions & {
    required?: boolean;
    caption?: string;
};

interface Block {
    blockType: BlockType;
    heading: string;
    id: string;
    paragraphs: Array<CustomParagraph>
}  

const blocks: Array<Block> = [{
        id: "localAuthSearchResult",
        blockType: "Normal",
        heading: "LOCAL AUTHORITY SEARCH RESULT",
        paragraphs: [
            {text: "I have made a search of the Register of Local Land Charges with the Local Authority and summarise below the information revealed:", required: true},
            {caption: "Planning charge", text: "Please see Part Three Planning Charge."}
        ],
    },
    {
        id: "anotherSearchResult",
        blockType: "Normal",
        heading: "ANOTHER SEARCH RESULT",
        paragraphs: [
            {text: "Another Paragraph"},
            {text: "And another!"}
        ],
    }, 
    {
        id: "laurasTest",
        blockType: "Normal",
        heading: "Laura's Test",
        paragraphs: [
            {text: "Another Paragraph", required: true},
            {text: "And another!", caption: "Some Caption"}
        ],
    }
];

const createReportHeader = (heading: Block["heading"], hasBreak: boolean) => 
    new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
            new TextRun({
                ...(hasBreak && {break: 1}),
                text: heading,
                bold: true,
                underline: {
                    color: "#ffffff",
                    type: "single"
                },
                size: 20
            })
        ]
    });

const createParagraphs = (paragraphs: Block["paragraphs"] ) =>
    new Paragraph({
        children: paragraphs.map((textRun, index) =>
            new TextRun({
                break: index === 0 ? 1 : 2,
                ...textRun
                }
            ),
        )
      });

const createReport = (report: Block, index: number): Array<Paragraph> => {
    const {heading, paragraphs } = report;
    return [
        createReportHeader(heading, index > 0),
        report.paragraphs.length > 0 ? createParagraphs(paragraphs) : new Paragraph({
            children: [
                new TextRun({
                    text: "You did not request this search.",
                    bold: true,
                    break: 1
                })
            ]
        })
    ];
}

const documentWidth = 9000;
const calculateStartXForCenter = (itemWidth: number) => (documentWidth/2) - (itemWidth/2);

const header = (ref: string, address: string) => [
    new Paragraph({
        frame: {
            type:"absolute",
            position: {
                x: 8000,
                y: 0,
            },
            width: 1500,
            height: 100,
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
        },
        children: [
            new TextRun({text: "Ref:", bold: true}),
            new TextRun(ref)
        ],
    }),
    new Paragraph({
        frame: {
            type:"absolute",
            position: {
                x: 0,
                y: 0,
            },
            width: documentWidth,
            height: 300,
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
        },
        children: [
            new ImageRun({
                data: imageBase64Data,
                transformation: {
                    width: 120,
                    height: 100,
                },
            }),
        ],
    }),
    new Paragraph({
        alignment: AlignmentType.CENTER,
        frame: {
            type:"absolute",
            position: {
                x: calculateStartXForCenter(documentWidth),
                y: 1200,
            },
            width: documentWidth,
            height: 100,
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
        },
        spacing: {
            line: 1500,
            lineRule: LineRuleType.AUTO,
        },
        children: [
            new TextRun({text: "SEARCH REPORT", bold: true, size: 25}),
        ],
    }),
    new Paragraph({
        alignment: AlignmentType.CENTER,
        shading: {
            type: ShadingType.SOLID,
            color: "DFDFDF",
        },
        frame: {
            type:"absolute",
            position: {
                x: 0,
                y: 2000,
            },
            width: documentWidth,
            height: 300,
            anchor: {
                horizontal: FrameAnchorType.MARGIN,
                vertical: FrameAnchorType.MARGIN,
            },
        },
        border: {
            top: {
                color: "auto",
                space: 10,
                style: "single",
                size: 6,
            },
            bottom: {
                color: "auto",
                space: 10,
                style: "single",
                size: 6,
            },
            left: {
                color: "auto",
                space: 1,
                style: "single",
                size: 6,
            },
            right: {
                color: "auto",
                space: 1,
                style: "single",
                size: 6,
            },
        },
        children: [
            new TextRun({text: "Address of the property", bold: true}),
            new TextRun({text: address, break: 1})
        ],
    }),
];

const searchReportSchema = object().shape({
    ref: string().required("Required"),
    address: string().required("Required")
})

export const DocsGenerator = () => {
    
   const [selectedBlocks, setSelectedBlocks] = useState<Block[]>(blocks.map(({paragraphs, ...rest}) => ({paragraphs: [], ...rest})));

   const handleChange = (event: SelectChangeEvent<string[]>, id: string) => {
     const {
       target: { value },
     } = event;
     setSelectedBlocks(currentBlocks => currentBlocks.map((block) => {
       if(block.id === id){
           return {
               ...block,
               paragraphs: value.length === 0 ? [] : blocks.find((b) => b.id === id)!.paragraphs.filter((para) => value.includes(para.text))
           }
       }
       return block;
     }));
   };

    const generate = ({ref, address}: any) => {
         const document = new Document({
            styles: {
                default: {
                    document: {
                        run: {
                            font: "Century Gothic"
                        },
                    }
                },
                paragraphStyles: [
                    {
                        id: "headerParagraph",
                        basedOn: "Normal",
                        run: {
                            shading: {
                                type: ShadingType.SOLID,
                                color: "00FFFF",
                                fill: "FF0000",
                            }
                        }
                    }
                ]
            },
            sections: [
                {
                    footers: {
                        default: new Footer({
                            children: [
                                new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                        new TextRun({
                                            children: ["RG Law - Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES],
                                        }),
                                    ],
                                }),
                            ],
                        }),
                    },
                    children: [
                        ...header(ref, address)
                    ]
                },
                {
                    properties: {
                        type: SectionType.CONTINUOUS,
                    },
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            frame: {
                                type:"absolute",
                                position: {
                                    x: 0,
                                    y: 3300,
                                },
                                width: documentWidth,
                                height: 300,
                                anchor: {
                                    horizontal: FrameAnchorType.MARGIN,
                                    vertical: FrameAnchorType.MARGIN,
                                },
                            },
                            children: [
                                new TextRun({
                                    text: "Please read and digest the following information in respect of the search results relating to your purchase of the  above property."
                                }),
                            ],
                        }),
                    ]
                },
                {
                    properties: {
                        type: SectionType.CONTINUOUS,
                    },
                    children: selectedBlocks.reduce<Array<Paragraph>>((acc, report, index) => [
                        ...acc, ...createReport(report, index)], [])
                }
            ]
         });

        Packer.toBlob(document).then(blob => {
          saveAs(blob, "example.docx");
        });
    }

    // Tidy, type, style and add success snackbar

    return (
        <Stack>
             <AppBar position="static">
                <Toolbar>
                    <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
                        Report Generator
                    </Typography>
                    <Button color="inherit">Search report</Button>
                    </Toolbar>
            </AppBar>
            <Formik 
                initialValues={{
                    ref: "",
                    address: ""
                }} 
                validationSchema={searchReportSchema}
                onSubmit={(values, {resetForm}) => {
                    generate(values);
                    resetForm();
                }}
                >
                    {({ isValid }) => (
                        <Form>     
                            <Stack alignItems="center" sx={{width: rem(500), m: 3}}>
                                <Button sx={{m: 1}} type="submit" variant="outlined" disabled={!isValid} onClick={generate}>Generate Report</Button>
                                <Field
                                    component={TextField}
                                    name="ref"
                                    type="text"
                                    label="Ref"
                                    style={{width: "100%", height: rem(80)}}
                                />
                                <Field
                                    component={TextField}
                                    name="address"
                                    type="text"
                                    label="Address"
                                    style={{width: "100%", height: rem(80)}}
                                />
                                
                                {blocks.map(({heading, paragraphs, id}) => 
                                    <FormControl key={id} sx={{ m: 1, width: "100%"}}>
                                        <InputLabel id={`${id}-checkbox-label`}>{heading}</InputLabel>
                                        <Select
                                            labelId={`${id}-checkbox-label`}
                                            id={`${id}-multiple-checkbox`}
                                            multiple
                                            value={selectedBlocks.find((sb) => sb.id === id)!.paragraphs.map(p => p.text)}
                                            onChange={(event) => handleChange(event, id)}
                                            input={<OutlinedInput label={heading} />}
                                            renderValue={(selected) => selected.join(', ')}
                                            >
                                            {paragraphs.map((para, index) => (
                                                <MenuItem key={index} value={para.text}>
                                                    <Checkbox checked={selectedBlocks.find((sb) => sb.id === id)!.paragraphs.indexOf(para) > -1} />
                                                    <Typography sx={{textWrap: "wrap"}}>{para.text}</Typography>
                                                </MenuItem>
                                            ))}
                                        </Select>
                                    </FormControl>
                                )}
                            </Stack>
                        </Form>
                    )}
            </Formik>   
        </Stack>
    );
}