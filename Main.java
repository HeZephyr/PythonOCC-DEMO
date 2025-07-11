import java.util.LinkedList;
import java.util.Queue;

class TreeNode {
    int val;
    TreeNode left;
    TreeNode right;
    TreeNode() {}
    TreeNode(int val) { this.val = val; }
    TreeNode(int val, TreeNode left, TreeNode right) {
        this.val = val;
        this.left = left;
        this.right = right;
    }
}

public class Main {
    public int minCameraCover(String input) {
        String[] nodes = input.split(" ");
        if (nodes.length == 0 || nodes[0].equals("N")) {
            return 0;
        }
        TreeNode root = buildTree(nodes);
        int[] res = dfs(root);
        return Math.min(res[1], res[2]);
    }

    private TreeNode buildTree(String[] nodes) {
        if (nodes.length == 0 || nodes[0].equals("N")) {
            return null;
        }
        TreeNode root = new TreeNode(Integer.parseInt(nodes[0]));
        Queue<TreeNode> queue = new LinkedList<>();
        queue.offer(root);
        int index = 1;
        while (!queue.isEmpty() && index < nodes.length) {
            TreeNode current = queue.poll();
            if (index < nodes.length && !nodes[index].equals("N")) {
                current.left = new TreeNode(Integer.parseInt(nodes[index]));
                queue.offer(current.left);
            }
            index++;
            if (index < nodes.length && !nodes[index].equals("N")) {
                current.right = new TreeNode(Integer.parseInt(nodes[index]));
                queue.offer(current.right);
            }
            index++;
        }
        return root;
    }

    private int[] dfs(TreeNode node) {
        if (node == null) {
            return new int[]{0, 0, 3001}; // state0, state1, state2
        }
        int[] left = dfs(node.left);
        int[] right = dfs(node.right);

        int state0 = Math.min(left[1], left[2]) + Math.min(right[1], right[2]);

        int option1 = left[2] + Math.min(right[1], right[2]);
        int option2 = right[2] + Math.min(left[1], left[2]);
        int option3 = left[2] + right[2];
        int state1 = Math.min(Math.min(option1, option2), option3);

        int leftMin = Math.min(left[0], Math.min(left[1], left[2]));
        int rightMin = Math.min(right[0], Math.min(right[1], right[2]));
        int state2 = 1 + leftMin + rightMin;

        return new int[]{state0, state1, state2};
    }
}