﻿using EasyOffice.Enums;
using EasyOffice.Models.Word;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace EasyOffice.Providers.NPOI
{
    public static class NPOIWordExtensions
    {
        //在NPOI中，每厘米对应的长度数值
        //private const int NPOI_PICTURE_LENGTH_EVERY_CM = 360144;
        private const int NPOI_PICTURE_LENGTH_EVERY_CM = 9525;
        public static XWPFDocument ToNPOI(this Word word)
        {
            XWPFDocument doc = null;
            using (MemoryStream ms = new MemoryStream(word.WordBytes))
            {
                doc = new XWPFDocument(ms);
            }

            return doc;
        }

        public static XWPFParagraph Set(this XWPFParagraph p, Paragraph paragraph)
        {
            p.Alignment = GetNPOIAlignment(paragraph.Alignment);
            p.CreateRun().Set(paragraph.Run);

            return p;
        }

        public static XWPFRun Set(this XWPFRun xwpfRun, Run run)
        {
            if (run != null)
            {
                if (run.HasLineBreak)
                {
                    xwpfRun.AddCarriageReturn();
                }
                else
                {
                    xwpfRun.SetText(run.Text);
                    xwpfRun.FontSize = run.FontSize;//设置字体大小
                    xwpfRun.SetFontFamily(run.FontFamily, FontCharRange.Ascii);//设置粗体
                    xwpfRun.IsBold = run.IsBold;
                    xwpfRun.SetColor(run.Color);
                    xwpfRun.IsItalic = run.IsItalic;
                    if (run.UnderlineType != UnderlineType.None)
                    {
                        switch (run.UnderlineType)
                        {
                            case UnderlineType.Underline:
                                xwpfRun.SetUnderline(UnderlinePatterns.Single);
                                break;
                            case UnderlineType.DoubleUnderline:
                                xwpfRun.SetUnderline(UnderlinePatterns.Double);
                                break;
                            case UnderlineType.ThickUnderline:
                                xwpfRun.SetUnderline(UnderlinePatterns.Thick);
                                break;
                            case UnderlineType.DottedDashed:
                                xwpfRun.SetUnderline(UnderlinePatterns.Dotted);
                                break;
                            case UnderlineType.Dashed:
                                xwpfRun.SetUnderline(UnderlinePatterns.Dash);
                                break;
                            case UnderlineType.WavyLine:
                                xwpfRun.SetUnderline(UnderlinePatterns.Wave);
                                break;
                            case UnderlineType.None:
                                xwpfRun.SetUnderline(UnderlinePatterns.None);
                                break;
                            default:
                                break;
                        }
                    }
                }

            }

            if (run.Pictures != null)
            {
                xwpfRun.Set(run.Pictures);
            }

            return xwpfRun;
        }

        public static XWPFRun Set(this XWPFRun xwpfRun, IEnumerable<Picture> pictures)
        {
            foreach (var picture in pictures)
            {
                var pictureData = picture.PictureData;
                if (pictureData == null || pictureData.Length == 0)
                {
                    try
                    {
                        pictureData = File.OpenRead(picture.PictureUrl);
                    }
                    catch (Exception)
                    {
                    }
                }

                if (pictureData == null || pictureData.Length == 0) continue;

                int height = (int)(Math.Ceiling(picture.Height * NPOI_PICTURE_LENGTH_EVERY_CM));
                int width = (int)(Math.Ceiling(picture.Width * NPOI_PICTURE_LENGTH_EVERY_CM));

                //int height = (int)(Math.Ceiling(picture.Height ));
                //int width = (int)(Math.Ceiling(picture.Width));
                xwpfRun.AddPicture(pictureData, picture.PictureType.GetHashCode(), picture.FileName, width, height);

                CT_Inline inline = xwpfRun.GetCTR().GetDrawingList()[0].inline[0];
                inline.docPr.id = 1;

                pictureData.Dispose();
            }

            return xwpfRun;
        }

        private static ParagraphAlignment GetNPOIAlignment(Alignment alignment)
        {
            switch (alignment)
            {
                case Alignment.LEFT:
                    return ParagraphAlignment.LEFT;
                case Alignment.CENTER:
                    return ParagraphAlignment.CENTER;
                case Alignment.RIGHT:
                    return ParagraphAlignment.RIGHT;
            }

            return ParagraphAlignment.CENTER;
        }

        public static byte[] ToBytes(this XWPFDocument doc)
        {
            byte[] result;
            using (MemoryStream ms = new MemoryStream())
            {
                doc.Write(ms);
                result = ms.ToArray();
            }

            return result;
        }
    }
}
