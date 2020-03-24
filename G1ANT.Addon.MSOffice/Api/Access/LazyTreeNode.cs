using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    internal class LazyTreeNode : TreeNode
    {
        private readonly List<Func<IEnumerable<TreeNode>>> treeNodeFactories;

        //public LazyTreeNode(string text, Func<IEnumerable<TreeNode>> treeNodeFactory) : base(text)
        //{
        //    this.treeNodeFactories = new Func<IEnumerable<TreeNode>>[] { treeNodeFactory };

        //    Nodes.Add("");
        //}

        //public LazyTreeNode(string text, params Func<IEnumerable<TreeNode>>[] treeNodeFactories) : base(text)
        //{
        //    this.treeNodeFactories = treeNodeFactories;

        //    Nodes.Add("");
        //}

        public LazyTreeNode(string text, params Func<IEnumerable<TreeNode>>[] treeNodeFactories) : base(text)
        {
            this.treeNodeFactories = treeNodeFactories.ToList();

            Nodes.Add("");
        }

        public LazyTreeNode(string text) : base(text)
        {
            this.treeNodeFactories = new List<Func<IEnumerable<TreeNode>>>();

            Nodes.Add("");
        }

        public LazyTreeNode Add(TreeNode treeNode)
        {
            treeNodeFactories.Add(() => new List<TreeNode>() { treeNode });
            return this;
        }

        public LazyTreeNode AddRange(params Func<IEnumerable<TreeNode>>[] treeNodeFactories)
        {
            this.treeNodeFactories.AddRange(treeNodeFactories);
            return this;
        }

    //public LazyTreeNode(string text, params Func<IEnumerable<LazyTreeNode>>[] treeNodeFactories) : base(text)
    //{
    //    this.treeNodeFactories = treeNodeFactories;

    //    Nodes.Add("");
    //}


    private bool IsEmpty()
        {
            return Nodes.Count == 1 && Nodes[0].Text == "";
        }

        public void LoadLazyChildren()
        {
            if (IsEmpty())
            {
                Nodes.Clear();
                try { treeNodeFactories.ToList().ForEach(f => Nodes.AddRange(f().ToArray())); }
                catch { }
            }
        }

    }
}
