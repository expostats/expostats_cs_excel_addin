using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpostatsExcel2013AddIn
{
class NDExpo
    {
        public class Datum
        {
            public double detectionLimitValue;
            public double value;
            public int position;
            public int index;
            public double plottingPosition;
            public double score;
            public double finalValue;
            public int isND = 0;
            public double predicted;
            public DetectionLimit dlGroup;
            
            public Datum(int isND, double value)
            {
                this.isND = isND; // 1 if not detected, 0 otherwise
                if (this.isND == 1)
                {
                    this.detectionLimitValue = value;
                }
                else
                {
                    this.value = value;
                }
                this.dlGroup = null;
                this.position = -1; // from 0 to n-1, where n is the size of the dataset ($nd.dataSet)
                this.index = -1;
                this.plottingPosition = -1;
                this.score = -1;
                this.predicted = -1;
                this.finalValue = -1;
            }
            public double getValue()
            {
                if (this.isNotDetected()) return this.detectionLimitValue; else return this.value;
            }

            public String toString()
            {
                return this.position + " : " + (this.isNotDetected() ? "<" : "") + this.getValue() + " : " + this.index + " : " + this.plottingPosition + " : " + this.score + " : " + this.finalValue;
            }

            public bool isNotDetected()
            {
                return this.isND == 1;
            }
        }; // End Datum

        public class DetectionLimit
        {
            public double value;
            public int A;
            public int B;
            public int C;
            public double overProbability;
            public double nextOverProbability;
            public List<double> previousValues;

            public DetectionLimit(double value)
            {
                this.value = value; // the value of the limit
                this.A = 0; // number of detected > this.value 
                this.B = 0; // number of datum where datum.getValue() < this.value
                this.C = 0; // number of datum such than datum.isNotDetected() && (datum.detectionLimitValue = thid.value)
                this.overProbability = 0; // probability that a value is over the limit
                this.nextOverProbability = 0;
                this.previousValues = new List<double>();
            }

            public double getPlottingPosition(Datum datum)
            {
                if (datum.isNotDetected())
                {
                    return datum.index * (1 - this.overProbability) / (this.C + 1);
                }
                else
                {
                    return (1 - this.overProbability) + (datum.index * ((this.overProbability - this.nextOverProbability) / (this.A + 1)));
                }
            }
        } // End DetectionLimit

        public class GlobalValues
        {
            public double sumX;
            public double sumY;
            public double sumXY;
            public double sumXX;
            public double slope;
            public double intercept;
            public int cnt;
            public double per20;
            public GlobalValues(List<Datum> dataSet)
            {
                Datum datum;
                int iDatum;
                double x, y;

                this.sumX = 0;
                this.sumY = 0;
                this.sumXY = 0;
                this.sumXX = 0;
                this.cnt = 0;

                this.per20 = -1; // undefined

                for (iDatum = 0; iDatum < dataSet.Count; iDatum++)
                {
                    datum = dataSet[iDatum];

                    if (!datum.isNotDetected())
                    {

                        y = Math.Log(datum.value);
                        x = datum.score;
                        this.sumX += x;
                        this.sumY += y;
                        this.sumXX += (x * x);
                        this.sumXY += (x * y);
                        this.cnt++;
                    }
                }

                this.slope = ((this.cnt * this.sumXY) - (this.sumX * this.sumY)) / ((this.cnt * this.sumXX) - (this.sumX * this.sumX));
                this.intercept = ((this.sumXX * this.sumY) - (this.sumX * this.sumXY)) / ((this.cnt * this.sumXX) - (this.sumX * this.sumX));
            }

            public double getPercentile(int pNTH, List<Datum> dataSet)
            {
                int pStep = dataSet.Count / 100;

                if (pNTH < 1)
                {
                    pNTH = 1;
                }

                int position = (int) Math.Round((double) pNTH * pStep) - 1;

                if (position < 0)
                {
                    position = 0;
                }
                else
                {
                    if (position >= dataSet.Count)
                    {
                        position = dataSet.Count - 1;
                    }
                }

                return dataSet[position].finalValue;

            }
        }

        public List<Datum> dataSet;
        int error = 0;
        bool grouping = true;
        public Count count = new Count(0, 0);

        public struct Count
        {
            public Count(int n, int t)
            {
                nd = n;
                total = t;
            }

            public int nd { get; set; }
            public int total { get; set; }

        }

        public const int ERR_TOOMANY_ND = 1;
        public const int ERR_GRTSTDL = 2;
        public const int ERR_NENGH_DATA = 3;
        public const int WARN_NO_ND = 4;
        public const int ERR_NENGH_DET = 5;

        public void reset() { // 
            this.dataSet = new List<Datum>();
            this.detectionLimitArray = new List<DetectionLimit>();
            this.count.nd = 0;
            this.count.total = 0;
            this.error = 0;
        }

        public int compareValue(Datum datum1, Datum datum2 )
        {
            double diff = datum1.getValue() - datum2.getValue();
            if (diff == 0)
            {
                diff = datum2.isND - datum1.isND;
            }

            int cmpVal = diff > 0 ? 1 : (diff < 0 ? -1 : 0);            
            return cmpVal;
        }

        public int compareFinalValue(Datum datum1, Datum datum2)
        {
            double diff = datum1.finalValue - datum2.finalValue;
            int cmpVal = diff > 0 ? 1 : (diff < 0 ? -1 : 0);
            return cmpVal;
        }

        public int comparePosition(Datum datum1, Datum datum2)
        { // restore initial order
            return datum1.position - datum2.position;
        }

        public void addDatum(int isND, double value, int? position )
        {
            var datum = new Datum(isND, value);
            datum.position = (position != null) ? (int) position : this.dataSet.Count();
            this.dataSet.Add(datum);
            this.count.total++;
            if (isND == 1)
            {
                this.count.nd++;
            }
        }

        void sortDataSet()
        {
            this.dataSet.Sort(this.compareValue);
        }

        public List<DetectionLimit> detectionLimitArray;

        public GlobalValues global;

        public void doCalc() {
            int iDatum, iLimit;
            Datum datum, prevDatum;
            int currentPosition;
            DetectionLimit currentLimit;
            double nextProb;
        
            this.dataSet.Sort(this.compareValue);
            int[] pos = this.dataSet.Select(d => d.position).ToArray();

            if (this.count.total < 5) {
                this.error = NDExpo.ERR_NENGH_DATA;
                return;
            }

            if ((this.count.total - this.count.nd) < 3)
            {
                this.error = NDExpo.ERR_NENGH_DET;
                return;
            }

            if (this.count.nd > Math.Floor(.8 * this.count.total))
            {
                this.error = NDExpo.ERR_TOOMANY_ND;
                return;
            }

            if (this.dataSet[this.dataSet.Count - 1].isNotDetected())
            {
                this.error = NDExpo.ERR_GRTSTDL;
                return;
            }
            if (this.count.nd == 0)
            {
                this.error = NDExpo.WARN_NO_ND;
            }
            this.detectionLimitArray = new List<DetectionLimit>();



            currentLimit = new DetectionLimit(0);
            this.detectionLimitArray.Add(currentLimit);

            datum = this.dataSet[0];
            if (datum.isND == 1)
            {
                currentLimit = new DetectionLimit(datum.detectionLimitValue);
                currentLimit.C++;
                this.detectionLimitArray.Add(currentLimit);
            }
            else
            {
                currentLimit.A++;
            }
            currentPosition = 1;
            datum.index = 1;
            datum.dlGroup = currentLimit;

            for (iDatum = 1; iDatum < this.dataSet.Count; iDatum++)
            {
                datum = this.dataSet[iDatum];
                prevDatum = this.dataSet[iDatum - 1];
                if (prevDatum.isND != datum.isND)
                {
                    if (datum.isND == 1)
                    {
                        currentLimit = new DetectionLimit(datum.detectionLimitValue);
                        this.detectionLimitArray.Add(currentLimit);
                    }
                    currentPosition = 0;
                }
                else
                {
                    if (datum.isND == 1 && (prevDatum.detectionLimitValue != datum.detectionLimitValue))
                    {
                        if (this.grouping)
                        {
                            currentLimit.previousValues.Add(currentLimit.value);
                            currentLimit.value = datum.detectionLimitValue;
                        }
                        else
                        {
                            currentLimit = new DetectionLimit(datum.detectionLimitValue);
                            this.detectionLimitArray.Add(currentLimit);
                            currentPosition = 0;
                        }

                    }
                }

                currentPosition++;
                datum.index = currentPosition;
                datum.dlGroup = currentLimit;
                if (datum.isND == 1)
                {
                    currentLimit.C++;
                }
                else
                {
                    currentLimit.A++;
                }

            }

            currentLimit = new DetectionLimit(this.dataSet[this.dataSet.Count - 1].value + 1);
            this.detectionLimitArray.Add(currentLimit);

            for (iLimit = 1; iLimit < this.detectionLimitArray.Count; iLimit++)
            {
                this.detectionLimitArray[iLimit].B = this.detectionLimitArray[iLimit].C + this.detectionLimitArray[iLimit - 1].B + this.detectionLimitArray[iLimit - 1].A;
            }


            for (iLimit = this.detectionLimitArray.Count - 2; iLimit >= 0; iLimit--)
            {
                currentLimit = this.detectionLimitArray[iLimit];
                nextProb = this.detectionLimitArray[iLimit + 1].overProbability;
                currentLimit.nextOverProbability = nextProb;
                currentLimit.overProbability = (double) nextProb + ((1 - nextProb) * (currentLimit.A / (double)  (currentLimit.A + currentLimit.B)));
            }

            this.detectionLimitArray[0].overProbability = 1; // this way, no rounding error

            for (iDatum = 0; iDatum < this.dataSet.Count; iDatum++)
            {
                datum = this.dataSet[iDatum];
                datum.plottingPosition = datum.dlGroup.getPlottingPosition(datum);
                datum.score = this.inverseNormalcdf(datum.plottingPosition);
            }

            this.global = new GlobalValues(this.dataSet);

            for (iDatum = 0; iDatum < this.dataSet.Count; iDatum++)
            {
                datum = this.dataSet[iDatum];
                if (datum.isNotDetected())
                {
                    datum.predicted = (datum.score * this.global.slope) + this.global.intercept;
                    datum.finalValue = Math.Exp(datum.predicted);
                }
                else
                {
                    datum.finalValue = datum.value;
                }
            }
            this.dataSet.Sort(this.compareFinalValue);
            pos = this.dataSet.Select(d => d.position).ToArray();
            this.global.per20 = this.global.getPercentile(20, this.dataSet);

            this.dataSet.Sort(this.comparePosition);
            pos = this.dataSet.Select(d => d.position).ToArray();
        }

        double inverseNormalcdf(double p) {
    
            double
            a1 = -39.6968302866538,
            a2 = 220.946098424521,
            a3 = -275.928510446969,
            a4 = 138.357751867269,
            a5 = -30.6647980661472,
            a6 = 2.50662827745924,

            b1 = -54.4760987982241,
            b2 = 161.585836858041,
            b3 = -155.698979859887,
            b4 = 66.8013118877197,
            b5 = -13.2806815528857,

            c1 = -7.78489400243029E-03,
            c2 = -0.322396458041136,
            c3 = -2.40075827716184,
            c4 = -2.54973253934373,
            c5 = 4.37466414146497,
            c6 = 2.93816398269878,

            d1 = 7.78469570904146E-03,
            d2 = 0.32246712907004,
            d3 = 2.445134137143,
            d4 = 3.75440866190742,

            //Define break-points
            p_low = 0.02425,
            p_high = 1 - p_low,

            //Define work variables
            q,
            r;

                // If argument out of bounds, raise error
                if (p <= 0 || p >= 1) throw new Exception();

            if (p < p_low)
            {
                //Rational approximation for lower region
                q = Math.Sqrt(-2 * Math.Log(p));
                return (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
            }
            else
                if (p <= p_high)
            {
                //Rational approximation for lower region
                q = p - 0.5;
                r = q * q;
                return (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1);
            }
            else
                    if (p < 1)
            {
                //Rational approximation for upper region
                q = Math.Sqrt(-2 * Math.Log(1 - p));
                return -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1);
            }
            return -99999;
        }

        public class GraphData
        {
            NDExpo ndObject;
            public List<double[]> nd { get; }
            public List<double[]> regression { get; }
            public List<double[]> detected { get; }
            double minX;
            double maxX;
            double minXDet;
            double maxXDet;
            double minY;
            double maxY;

            public GraphData(NDExpo ndo)
            {
                this.ndObject = ndo;
                this.nd = new List<double[]>();
                this.regression = new List<double[]>();
                this.detected = new List<double[]>();
                this.minX = double.PositiveInfinity;
                this.maxX = double.NegativeInfinity;
                this.minXDet = double.PositiveInfinity;
                this.maxXDet = double.NegativeInfinity;
                this.minY = double.PositiveInfinity;
                this.maxY = double.NegativeInfinity;
            }

            public void getDataForChart()
            {
                List<Datum> data = this.ndObject.dataSet;
                int iDatum;
                Datum datum;
                double y;
                double[] xyPoint;

                for (iDatum = 0; iDatum < data.Count; iDatum++)
                {
                    datum = data[iDatum];
                    y = Math.Log(datum.finalValue);

                    xyPoint = new double[]{ datum.score, y, datum.finalValue };
                    this.minX = Math.Min(this.minX, xyPoint[0]);
                    this.maxX = Math.Max(this.maxX, xyPoint[0]);
                    this.minY = Math.Min(this.minY, y);
                    this.maxY = Math.Max(this.maxY, y);

                    if (datum.isNotDetected())
                    {
                        this.nd.Add(xyPoint);
                    }
                    // le else est inutile. On trace la droite pour tous les points
                    else
                    {
                        this.detected.Add(xyPoint);
                        this.minXDet = Math.Min(this.minXDet, xyPoint[0]);
                        this.maxXDet = Math.Max(this.maxXDet, xyPoint[0]);
                    }
                }
                // On utilisait 'minXDet et maxXDet
                y = this.ndObject.global.slope * this.minX + this.ndObject.global.intercept;
                this.regression.Add( new double[] { this.minX, y, Math.Exp(y) } );
                y = this.ndObject.global.slope * this.maxX + this.ndObject.global.intercept;
                this.regression.Add( new double[] { this.maxX, y, Math.Exp(y) } );
                this.minY = Math.Min(this.minY, this.regression[0][1]);
                this.maxY = Math.Max(this.maxY, this.regression[1][1]);
            }
        }
    }
}
