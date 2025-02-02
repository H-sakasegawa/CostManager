using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using ExcelReaderUtility;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using static CostManager.MaterialCost;

namespace CostManager
{
    [Serializable]
    public class CostDataList
    {
        public CostDataList() { }

        public void Add(CostData cost)
        {
            ListCostDatas.Add(cost);
        }
        public List<CostData> ListCostDatas { get; set; } = new List<CostData>();
    }



    [Serializable]
    public class CostData
    {

        /// <summary>
        /// 商品コード
        /// </summary>
        public string ProductId { get; set; }
        /// <summary>
        /// 商品分類
        /// </summary>
        public string Kind { get; set; }
        /// <summary>
        /// 商品名称
        /// </summary>
        public string ProductName { get; set; }
        /// <summary>
        /// 制作数
        /// </summary>
        public uint ProductNum { get; set; } = 0;
        /// <summary>
        /// 定価
        public uint Price { get; set; } = 0;

        /// <summary>
        /// 原材料 一覧
        /// </summary>
        public List<MaterialCost> LstMaterialCost { get; set; } = new List<MaterialCost>();
        /// <summary>
        /// 人件費 一覧
        /// </summary>
        public List<WorkerCost> LstWorkerCost { get; set; } = new List<WorkerCost>();
        /// <summary>
        /// 包装材 一覧
        /// </summary>
        public List<PackageCost> LstPackageCost { get; set; } = new List<PackageCost>();

        public CostData() { }

        public CostData(ProductReader.ProductData product, CostReader costReader)
        {
            ProductId = product.id;
            ProductName = Utility.RemoveCRLF(product.name);
            Kind = Utility.RemoveCRLF(product.kind);

#if DEBUG
            ProductNum = 50;
            Price = 1500;

            LstMaterialCost.Add(new MaterialCost("1", 10, costReader));
            LstMaterialCost.Add(new MaterialCost("2", 22, costReader));
            LstMaterialCost.Add(new MaterialCost("3", 33, costReader));
            LstMaterialCost.Add(new MaterialCost("4", 44, costReader));

            LstWorkerCost.Add(new WorkerCost("1", 0, costReader));
            LstWorkerCost.Add(new WorkerCost("2", 0, costReader));

            LstPackageCost.Add(new PackageCost("1", ProductNum, costReader));
            LstPackageCost.Add(new PackageCost("2", ProductNum, costReader));
            LstPackageCost.Add(new PackageCost("3", ProductNum, costReader));
            LstPackageCost.Add(new PackageCost("4", ProductNum, costReader));
            LstPackageCost.Add(new PackageCost("5", ProductNum, costReader));
#else
            //以下の作業者と、包装材はデフォルト設定とする
            LstWorkerCost.Add(new WorkerCost("1", 0, costReader));
            LstWorkerCost.Add(new WorkerCost("2", 0, costReader));

            LstPackageCost.Add(new PackageCost("1", 1, costReader));
            LstPackageCost.Add(new PackageCost("2", 1, costReader));
            LstPackageCost.Add(new PackageCost("3", 1, costReader));
            LstPackageCost.Add(new PackageCost("4", 1, costReader));
            LstPackageCost.Add(new PackageCost("5", 1, costReader));

#endif


        }

        public void AttachCostReader(CostReader costBaseInfo)
        {
            foreach (var val in LstMaterialCost)
            {
                val.AttachCostReader(costBaseInfo);
            }
            foreach (var val in LstWorkerCost)
            {
                val.AttachCostReader(costBaseInfo);
            }
            foreach (var val in LstPackageCost)
            {
                val.AttachCostReader(costBaseInfo);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="costRate">原価率(%)</param>
        /// <param name="rowCostRate">原材料(%)</param>
        /// <param name="laborCostRate">人件費(%)</param>
        /// <param name="packageCostRate">包装費(%)</param>
        /// <param name="profit">利益</param>
        /// <returns> 0..正常 -1..原価割れ</returns>
        public int Calc(out float costRate, out float rowCostRate, out float laborCostRate, out float packageCostRate, out float profit)
        {
            costRate = 0;
            rowCostRate = 0;
            laborCostRate = 0;
            packageCostRate = 0;
            profit = 0;

            //n個制作時の原材料費
            var rowCost = GetMateralCost();
            //n個制作時の人件費
            var laborCost = GetWorkerCost();
            //n個制作時の包装費
            var packageCost = GetPackageCost();

            //原価
            var cost = rowCost + laborCost + packageCost;
            if (Price < cost)
            {
                //原価割れ
                return -1;
            }

            //各原価率
            if (Price != 0)
            {
                rowCostRate = (float)rowCost / Price;
                laborCostRate = (float)laborCost / Price;
                packageCostRate = (float)packageCost / Price;
            }


            //原価率
            costRate = (rowCostRate + laborCostRate + packageCostRate);

            //利益（n個販売時の定価 - n個販売時の定価定価　 *原価率)
            profit = Price - Price * costRate;

            return 0;
        }

        public float GetMateralCost()
        {
            return LstMaterialCost.Sum(x => x.CalcCost());
        }
        public float GetWorkerCost()
        {
            return LstWorkerCost.Sum(x => x.CalcCost());
        }
        public float GetPackageCost()
        {
            return LstPackageCost.Sum(x => x.CalcCost(ProductNum));
        }
        public float GetAllCost()
        {
            return GetMateralCost() + GetWorkerCost() + GetPackageCost();
        }
        //単価
        public float GetCostOne()
        {
            //製品単価
            if (ProductNum <= 0)
            {
                return 0;
            }
            return (GetAllCost() / ProductNum);
        }
        /// <summary>
        /// 原価率
        /// </summary>
        /// <returns></returns>
        public float GetCostRate()
        {
            var rowCost = GetMateralCost();
            var laborCost = GetWorkerCost();
            var packageCost = GetPackageCost();

            //各原価率
            var rowCostRate = (Price != 0 ? (float)rowCost / Price : 0);
            var laborCostRate = (Price != 0 ? (float)laborCost / Price : 0);
            var packageCostRate = (Price != 0 ? (float)packageCost / Price : 0);
            //原価率
            return (rowCostRate + laborCostRate + packageCostRate);
        }
        /// <summary>
        /// 利益
        /// </summary>
        /// <returns></returns>
        public float GetProfit()
        {
            var costRate = GetCostRate();
            //利益（n個販売時の定価 - n個販売時の定価定価　 *原価率)
            return Price - Price * costRate;
        }
        /// <summary>
        /// 利益率
        /// </summary>
        /// <returns></returns>
        public float GetProfitRate()
        {
            var profit = GetProfit();
            //利益（n個販売時の定価 - n個販売時の定価定価　 *原価率)
            return profit / Price;
        }
        //原価割れ判定
        public bool IsBelowCost()
        {
            return GetAllCost() > Price ? true : false;
        }

        public void CopyTo(CostData data)
        {
            data.ProductId = ProductId;
            data.Kind = Kind;
            data.ProductName = ProductName;
            data.ProductNum = ProductNum;
            data.Price = Price;
            data.LstMaterialCost.Clear();
            data.LstMaterialCost.AddRange(LstMaterialCost);
            data.LstWorkerCost.Clear();
            data.LstWorkerCost.AddRange(LstWorkerCost);
            data.LstPackageCost.Clear();
            data.LstPackageCost.AddRange(LstPackageCost);
        }
    }

    [Serializable]
    public class CostBase
    {
        public CostBase() { }
        public CostBase(CostReader costBaseInfo)
        {
            this.costBaseInfo = costBaseInfo;
        }

        public void AttachCostReader(CostReader costBaseInfo)
        {
            this.costBaseInfo = costBaseInfo;
        }

        [NonSerialized]
        protected CostReader costBaseInfo;
    }

    /// <summary>
    /// 原材料
    /// </summary>
    [Serializable]
    public class MaterialCost : CostBase
    {
        /// <summary>
        /// 原材料ID
        /// </summary>
        public string ID { get; set; } = null;
        /// <summary>
        /// 使用量(g)
        /// </summary>
        public uint AmountUsed { get; set; } = 0;

        public MaterialCost() { }
        public MaterialCost(string materialId, uint amountUsed, CostReader costBaseInfo) :
            base(costBaseInfo)
        {
            this.ID = materialId;
            this.AmountUsed = amountUsed;
        }
        /// <summary>
        /// 費用計算
        /// </summary>
        /// <returns></returns>
        public uint CalcCost()
        {
            //原材料IDから原材料情報の原価を取得 ★★
            var value = costBaseInfo.GetMaterialtDataByID(ID);
            if (value == null) return 0;
            return value.cost * AmountUsed;
        }
        public string GetKind()
        {
            var value = costBaseInfo.GetMaterialtDataByID(ID);
            if (value == null) return null;
            return value.kind;
        }
        public string GetName()
        {
            var value = costBaseInfo.GetMaterialtDataByID(ID);
            if (value == null) return null;
            return value.name;
        }
        public uint GetUsedAmount()
        {
            var value = costBaseInfo.GetMaterialtDataByID(ID);
            if (value == null) return 0;
            return value.gram;
        }
        /// <summary>
        /// 原材料１個辺りの原価
        /// </summary>
        public uint GetCost()
        {
            var value = costBaseInfo.GetMaterialtDataByID(ID);
            if (value == null) return 0;
            return value.cost;
        }
}

    /// <summary>
    /// 人件費
    /// </summary>
    [Serializable]
    public class WorkerCost : CostBase
    {
        public WorkerCost() { }
        /// <summary>
        /// 作業者ID
        /// </summary>
        public string ID { get; set; } = null;
        /// <summary>
        /// 作業時間（分）
        /// </summary>
        public uint WorkingTime { get; set; } = 0;

        public WorkerCost(string id, uint workingTime, CostReader costBaseInfo) :
            base(costBaseInfo)
        {
            this.ID = id;
            //this.name = name;
            this.WorkingTime = workingTime;
        }
        /// <summary>
        /// 費用計算
        /// </summary>
        /// <returns></returns>
        public float CalcCost()
        {
            //作業者から作業者情報の時給を取得
            var value = costBaseInfo.GetWorkerDataByID(ID);
            if (value == null) return 0;
            float costPerMin = (float)value.hourlyPay / 60;   //分給
            return (float)WorkingTime * costPerMin; //作業時間(分) ×1分当たりの費用
        }
        public string GetName()
        {
            var value = costBaseInfo.GetWorkerDataByID(ID);
            if (value == null) return null;
            return value.name;
        }

    }
    /// <summary>
    /// 包装材
    /// </summary>
    [Serializable]
    public class PackageCost : CostBase
    {
        public PackageCost() { }
        /// <summary>
        /// 包装材ID
        /// </summary>
        public string ID { get; set; }
        /// <summary>
        /// 必要数
        /// </summary>
        public uint RequiredNum { get; set; } = 0;


        public PackageCost(string id, uint requiredNum, CostReader costBaseInfo) :
            base(costBaseInfo)
        {
            this.ID = id;
            //this.name = name;
            this.RequiredNum = requiredNum;
        }
        /// <summary>
        /// 費用計算
        /// </summary>
        /// <returns></returns>
        public float CalcCost(uint num)
        {
            //包装材から包装材情報の費用を取得
            var value = costBaseInfo.GetPackageDataByID(ID);
            if (value == null) return 0;

            return value.cost * num;
        }
        public float CalcCost()
        {
            //包装材から包装材情報の費用を取得
            return CalcCost(RequiredNum);
        }
        public string GetName()
        {
            var value = costBaseInfo.GetPackageDataByID(ID);
            if (value == null) return null;
            return value.name;
        }    }
}
