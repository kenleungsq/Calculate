using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Calculate
{
    public class ExcelFunction
    {
        
        List<Rank> LiKeRanks = new List<Rank>();
        List<Rank> WenKeRanks = new List<Rank>();

        static string publicPath = @"L:\19投档分\19投档分\【已整理】8.23上海\最终";

        //生成文件路径
        string ResultPath = publicPath + @"\1.xlsx";

        //一开始的原始文件路径
        string YuanShiPath = publicPath + @"\shanghai1.xlsx";

        //一分一段表路径
        string RankFilePath = publicPath + @"\上海一分一段表（已检查）.xlsx";

        //文科录取人数
        public int wenKePersons = 50000;
        //理科录取人数
        public int liKePersons = 50001;

        void ReadScoreRank()
        {
            FileStream file = new FileStream(RankFilePath, FileMode.Open);
            XSSFWorkbook wb = new XSSFWorkbook(file);

            //选择文科sheet和理科sheet
            List<XSSFSheet> sheets = new List<XSSFSheet>();
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                sheets.Add((XSSFSheet)wb.GetSheetAt(i));
            }

            foreach (var sheet in sheets)
            {
                var rows = sheet.GetRowEnumerator();
                ICell scoreCell = null;
                ICell rankCell = null;
                ICell keLeiCell = null;



                while (rows.MoveNext())
                {
                    var row = (XSSFRow)rows.Current;

                    for (int i = 0; i < row.LastCellNum; i++)
                    {
                        if (row.RowNum == 0)
                        {
                            switch (row.GetCell(i).StringCellValue)
                            {
                                case "score":
                                    scoreCell = row.GetCell(i);
                                    break;
                                case "total":
                                    rankCell = row.GetCell(i);
                                    break;
                                case "科类":
                                    keLeiCell = row.GetCell(i);
                                    break;
                            }
                        }
                    }

                    if (row.RowNum != 0)
                    {
                        //取得每一行的最新值
                        scoreCell = row.GetCell(scoreCell.ColumnIndex);
                        if (scoreCell == null || scoreCell.CellType == CellType.Blank)
                        {
                            break;
                        }

                        rankCell = row.GetCell(rankCell.ColumnIndex);
                        keLeiCell = row.GetCell(keLeiCell.ColumnIndex);

                        if (scoreCell.CellType == CellType.Numeric && keLeiCell.StringCellValue.Contains("文"))
                        {
                            WenKeRanks.Add(new Rank(scoreCell.NumericCellValue, rankCell.NumericCellValue, keLeiCell.StringCellValue));
                        }
                        else if (scoreCell.CellType == CellType.Numeric && keLeiCell.StringCellValue.Contains("理"))
                        {
                            LiKeRanks.Add(new Rank(scoreCell.NumericCellValue, rankCell.NumericCellValue, keLeiCell.StringCellValue));
                        }

                    }

                }
            }
           



            wb.Close();
            file.Close();

        }
        double SelectRank(double highRank,double keLei)
        {
            if (WenKeRanks.Count == 0 || LiKeRanks.Count == 0)
            {
                ReadScoreRank();
            }

            List<Rank> ranks = new List<Rank>();
            if (keLei == 1)
            {
                ranks = WenKeRanks;
            }
            else if (keLei == 2)
            {
                ranks = LiKeRanks;
            }

            //不可超过最高排名，否则报错
            var minRank = ranks.Min(t => t.Rank1);
            //选出最接近的最高排名分数
            if (minRank <= highRank)
            {
                var a = from t in ranks
                    where t.Rank1 <= highRank
                    select t.Score;

                return a.Last();
            }
            else
            {
                //超过最高排名则返回最高分
                return ranks.Max(t => t.Score);
            }
        }
        public void ReadExcel()
        {
            FileStream file = new FileStream(YuanShiPath, FileMode.Open,FileAccess.ReadWrite);
            FileStream resultFile = new FileStream(ResultPath, FileMode.Create);
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = (XSSFSheet) wb.GetSheetAt(0);
            var rows = sheet.GetRowEnumerator();

            //创建新行
            rows.MoveNext();
            var row = (XSSFRow)rows.Current;
            ICell temCell = row.CreateCell(row.LastCellNum);
            temCell.SetCellValue("tem");
            ICell highRankCell = row.CreateCell(row.LastCellNum);
            highRankCell.SetCellValue("模拟最高分排位");
            ICell highScore = row.CreateCell(row.LastCellNum);
            highScore.SetCellValue("模拟最高分");
            ICell averageScore = row.CreateCell(row.LastCellNum);
            averageScore.SetCellValue("模拟平均分");

            ICell lowRankCell = row.GetCell(0);
            ICell lowScoreCell = row.GetCell(0);
            ICell keLeiCell = row.GetCell(0);
            for (int i = 0; i < row.LastCellNum; i++)
            {
                //获取最低分排名列序号
                if (row.GetCell(i).StringCellValue.Contains("最低分排名"))
                {
                    lowRankCell = row.GetCell(i);
                    
                }

                //获取最低分列序号
                if (row.GetCell(i).StringCellValue.Contains("投档分") || row.GetCell(i).StringCellValue.Equals("最低分"))
                {
                    lowScoreCell = row.GetCell(i);

                }

                //获取科类列序号
                if (row.GetCell(i).StringCellValue.Contains("科类"))
                {
                    keLeiCell = row.GetCell(i);
                }
            }

            while (rows.MoveNext())
            {
                row = (XSSFRow) rows.Current;

                //获取每行最低分排位、科类、最低分单元格
                lowRankCell = row.GetCell(lowRankCell.ColumnIndex);
                keLeiCell = row.GetCell(keLeiCell.ColumnIndex);
                lowScoreCell = row.GetCell(lowScoreCell.ColumnIndex);

                //创建每行的tem、模拟最高分排位、模拟最高分、模拟平均分单元格
                temCell = row.CreateCell(temCell.ColumnIndex);
                highRankCell = row.CreateCell(highRankCell.ColumnIndex);
                highScore = row.CreateCell(highScore.ColumnIndex);
                averageScore = row.CreateCell(averageScore.ColumnIndex);

                //有时空行也会继续向下
                if (lowRankCell == null && keLeiCell == null && lowScoreCell == null)
                {
                    break;
                }


                if (lowRankCell.CellType == CellType.Numeric && keLeiCell.NumericCellValue == 1)
                {
                    temCell.SetCellValue(lowRankCell.NumericCellValue / (1 + lowRankCell.NumericCellValue * 6 / wenKePersons));
                    highRankCell.SetCellValue(lowRankCell.NumericCellValue - temCell.NumericCellValue);
                    highScore.SetCellValue(SelectRank(highRankCell.NumericCellValue, keLeiCell.NumericCellValue));
                    averageScore.SetCellValue((highScore.NumericCellValue + lowScoreCell.NumericCellValue)/2.001);
                }
                else if (lowRankCell.CellType == CellType.Numeric && keLeiCell.NumericCellValue == 2)
                {
                    temCell.SetCellValue(lowRankCell.NumericCellValue / (1 + lowRankCell.NumericCellValue * 6 / liKePersons));
                    highRankCell.SetCellValue(lowRankCell.NumericCellValue - temCell.NumericCellValue);
                    highScore.SetCellValue(SelectRank(highRankCell.NumericCellValue, keLeiCell.NumericCellValue));
                    averageScore.SetCellValue((highScore.NumericCellValue + lowScoreCell.NumericCellValue) / 2.001);
                }


                
            }

            wb.Write(resultFile);
            wb.Close();
            //file.Flush();
            file.Close();

        }
        

    }
}