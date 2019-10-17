namespace Calculate
{
    public class Rank
    {
        public double Score;
        public double Rank1;
        public string KeLei;

        public Rank(double score, double rank, string keLei)
        {
            Score = score;
            Rank1 = rank;
            KeLei = keLei;
        }
    }
}