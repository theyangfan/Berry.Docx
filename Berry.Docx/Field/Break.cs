﻿using System;
using System.Collections.Generic;
using System.Text;

using O = DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Berry.Docx.Field
{
    public class Break : ParagraphItem
    {
        #region Private Members
        private readonly W.Break _break;
        #endregion

        #region Constructors
        public Break(Document doc, BreakType type) : this(doc, new W.Run(), new W.Break())
        {
            Type = type;
        }

        internal Break(Document doc, W.Run ownerRun, W.Break br) 
            : base(doc, ownerRun, br)
        {
            _break = br;
        }
        #endregion

        #region Public Properties
        public override DocumentObjectType DocumentObjectType => DocumentObjectType.Break;

        public BreakType Type
        {
            get
            {
                if (_break.Type == null) return BreakType.TextWrapping;
                return _break.Type.Value.Convert<BreakType>();
            }
            set
            {
                if (value == BreakType.TextWrapping)
                    _break.Type = null;
                else
                    _break.Type = value.Convert<W.BreakValues>();
            }
        }

        public BreakTextRestartLocation Clear
        {
            get
            {
                if(Type != BreakType.TextWrapping) return BreakTextRestartLocation.None;
                if (_break.Clear == null) return BreakTextRestartLocation.None;
                return _break.Clear.Value.Convert<BreakTextRestartLocation>();
            }
            set
            {
                if(value == BreakTextRestartLocation.None)
                    _break.Clear = null;
                else
                    _break.Clear = value.Convert<W.BreakTextRestartLocationValues>();
            }
        }
        #endregion
    }
}