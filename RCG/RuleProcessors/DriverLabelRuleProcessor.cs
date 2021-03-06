﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace RCG
{
    public class DriverLabelRuleProcessor : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        private bool IsNetAddress(string path)
        {
            return Utility.GetLocationType(path) == LocationType.Network;
        }

        private string GetNetAddressCategory(string path)
        {
            Match m = Regex.Match(path, @"^\\\\\d{1,3}?\.\d{1,3}\.\d{1,3}\.\d{1,3}\\(.*)\\");
            if (m.Groups.Count == 1)
                return m.Groups[0].ToString().Trim();
            else
                return m.Groups[m.Groups.Count - 1].ToString().Trim();
        }

        public override string Process(string source)
        {
            base.PreProcess(source);

            if (IsNetAddress(source))
                return GetNetAddressCategory(source);
            else
            {
                DriveInfo drive = new DriveInfo(Path.GetPathRoot(source));
                return drive.VolumeLabel;
            }
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new DriverLabelRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private DriverLabelRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }
    }
}
