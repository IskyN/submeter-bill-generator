import matplotlib, decimal

matplotlib.use("Agg")

from matplotlib import pyplot as plt


def autolabel(chart):
    """
    Attach a text label above each bar displaying its height
    """
    for rect in chart:
        height = rect.get_height()
        if height > 0:
            ax.text(rect.get_x() + rect.get_width() / 2., height + 0.1,
                    str(round(height, 2)), ha='center', va='bottom')

l = [0,1,decimal.Decimal(2.5),3]
r = range(len(l))
fig, ax = plt.subplots()

ax.set_xticks(r)
ax.set_xticklabels(('G1\nhi', 'G2\nbye', 'G3', 'G4'))
ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=90)
ax.axes.get_yaxis().set_visible(False)
ax.spines['top'].set_visible(False)
ax.spines['left'].set_visible(False)
ax.spines['right'].set_visible(False)
graph = ax.bar(r, l, 1/1.5, color="blue")
graph[2].set_color('red')
autolabel(graph)
fig.savefig("testingtesting124.png")
plt.close(fig)


