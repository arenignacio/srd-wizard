import * as fs from "fs";
import { AlignmentType, Document, Packer, Paragraph, TextRun, convertInchesToTwip, HeadingLevel, NumberFormat, Footer, PageNumber, Header, ImageRun, HorizontalPosition, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, PageOrientation } from "docx";
import  fitImage from "./utility.js";
import imageSize from "image-size";
import sharp from "sharp";

//PROPS

const images = {
    rfsLogo: fs.readFileSync("./assets/rfs.png"),
    banner: fs.readFileSync("./assets/banner.jpg"),
    nsLogo: fs.readFileSync("./assets/ns.png"),
    suiteapp: fs.readFileSync("./assets/suiteapps.png")
}

const srd = {
    customerName: 'Atlanta Wheels',
    functionName: 'Sales Order Picking'
}

const Title = srd.customerName + ' - ' + srd.functionName



// Documents contain sections, you can have multiple sections per document,
const doc = new Document({
    title: Title,
    creator: 'SRD Wizard',
    description: '',
    sections: [    
        //FRONT PAGE
        {
           
            properties: {
                titlePage: true,                
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.NUMBER_IN_DASH
                    },
                    margin: {
                        left: 0,
                        right: 0,
                        top: 5
                    },
                    size: {
                        orientation: PageOrientation.LANDSCAPE
                    }
                }
            },

            headers: {
                first: new Header({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [                                 
                                new ImageRun({
                                    type: 'png',
                                    data: images.rfsLogo,
                                    transformation: {
                                        width: imageSize(images.rfsLogo).width * .8,
                                        height: imageSize(images.rfsLogo).height * .8
                                    },
                                }),
                                new TextRun({
                                    text: "\tYour Supply Chain: Mobile. Accurate. Real-Time.",
                                    size: '14pt',
                                    color: '000000',
                                    font: 'Myriad Pro'
                                }),
                               
                            ]
                        })
                    ]
                })
            },

            children: [

                //Banner
                //Would not load, possibly not enough permission?
                /* new ImageRun({
                    type: 'jpg',
                    data: images.banner,
                    transformation: {
                        width: imageSize(images.banner).width,
                        height: imageSize(images.banner).height
                    },
                }), */

                //Title
                new Paragraph({
                    alignment: AlignmentType.CENTER,                   
                    children: [
                        new TextRun({
                            text: 'RF-SMART SYSTEMS REQUIREMENT DOCUMENT',
                            break: 10,
                            font: 'Lato Bold',
                            size: '16pt',
                            bold: true,
                            color: '000000'
                        }),
                        new TextRun({
                            text: '(SRD)',
                            break: 1,
                            font: 'Lato Bold',
                            size: '16pt',
                            bold: true,
                            color: '000000'
                        }),
                    ],

                }),
                new Paragraph({  
                    alignment: AlignmentType.CENTER,
                    children: [
                        new ImageRun({
                            type: 'png',
                            data: images.nsLogo,
                            transformation: {
                                width: imageSize(images.rfsLogo).width * .7,
                                height: imageSize(images.rfsLogo).height * .6
                            },
                            floating: {
                                horizontalPosition: {
                                    offset: _moveByUnits(70)
                                },
                                verticalPosition: {
                                    offset: _moveByUnits(435)
                                }
                            }
                        }),                        
                        new ImageRun({
                            type: 'png',
                            data: images.suiteapp,
                            transformation: {
                                width: 250,
                                height: 148
                            },
                            floating: {
                                horizontalPosition: {
                                    offset: _moveByUnits(750),
                                },
                                verticalPosition: {
                                    offset: _moveByUnits(430)
                                }
                            }
                        }),
                        new TextRun({
                            text: 'Function: ' + srd.functionName, 
                            break: 2,
                            font: 'Lato',
                            size: '12pt'
                        }),
                        new TextRun({
                            text: 'Customer: ' + srd.customerName,
                            break: 2,
                            font: 'Lato',
                            size: '12pt'
                        })
                            
                    ],
                }),
            ],
            footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({
                                    children: [
                                        PageNumber.CURRENT
                                    ]
                                })
                            ]
                        })
                    ]
                })
            },
        },
        {
            properties: {
                
                page: {
                    pageNumbers: {
                        formatType: NumberFormat.NUMBER_IN_DASH
                    },
                    size: {
                        orientation: PageOrientation.LANDSCAPE
                    }
                }
                
            },
            
           children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Customer: ' + srd.customerName,
                            break: 2,
                            font: 'Lato',
                            size: '12pt'
                        })
                    ]
                })
           ],

           footers: {
                default: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                                new TextRun({
                                    children: [
                                        PageNumber.CURRENT
                                    ]
                                })
                            ]
                        })
                    ]
                })
        },
        }
    ],
});

//helper functions
function _moveByUnits (num) {
    
    return num * 10000
}

// Used to export the file into a .docx file
Packer.toBuffer(doc).then((buffer) => {

    try {
        fs.writeFileSync(Title + '.docx', buffer);
        console.log('word doc sucessfully created')
    } catch (e) {
        console.log(e.message)
    }
});