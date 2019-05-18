using EasyOffice;
using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests.Models
{
    /// <summary>
    /// 日程日期
    /// </summary>
    public class ScheduleDate
    {
        public string DateTimeStr { get; set; }
        public List<Address> Addresses { get; set; }
    }

    /// <summary>
    /// 会场
    /// </summary>
    public class Address
    {
        public string Name { get; set; }
        public List<Category> Categories { get; set; }
    }

    /// <summary>
    /// 日程种类
    /// </summary>
    public class Category
    {
        public string Name { get; set; }
        public List<Schedule> Schedules { get; set; }
    }

    /// <summary>
    /// 日程
    /// </summary>
    public class Schedule
    {
        public string Name { get; set; }
        public string TimeString { get; set; }
        public List<Speaker> Speakers { get; set; }
    }

    /// <summary>
    /// 嘉宾
    /// </summary>
    public class Speaker
    {
        public string Position { get; set; }
        public string Name { get; set; }
    }
}
