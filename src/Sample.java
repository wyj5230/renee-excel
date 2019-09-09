public class Sample
{
    String id;
    String yearSeason;
    String event;

    public Sample(String id, String yearSeason, String event)
    {
        this.id = id;
        this.yearSeason = yearSeason;
        this.event = event;
    }

    public String getId()
    {
        return id;
    }

    public void setId(String id)
    {
        this.id = id;
    }

    public String getYearSeason()
    {
        return yearSeason;
    }

    public void setYearSeason(String yearSeason)
    {
        this.yearSeason = yearSeason;
    }

    public String getEvent()
    {
        return event;
    }

    public void setEvent(String event)
    {
        this.event = event;
    }
}
