public class Subject
{
    public Subject(Group group) 
    {
        Groups.Add(group);
    }
    public List<Group> Groups { get; set; } = new List<Group>();
}
