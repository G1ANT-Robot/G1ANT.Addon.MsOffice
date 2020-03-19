using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    internal class LazyTreeNode : TreeNode
    {
        private readonly Func<IEnumerable<TreeNode>> treeNodeFactory;

        public LazyTreeNode(string text, Func<IEnumerable<TreeNode>> treeNodeFactory) : base(text)
        {
            this.treeNodeFactory = treeNodeFactory;

            Nodes.Add("");
        }

        private bool IsEmpty()
        {
            return Nodes.Count == 1 && Nodes[0].Text == "";
        }

        public void LoadLazyChildren()
        {
            if (IsEmpty())
            {
                Nodes.Clear();
                Nodes.AddRange(treeNodeFactory().ToArray());
            }
        }

    }
}
