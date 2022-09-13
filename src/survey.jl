using XLSX, DataFrames, CairoMakie, HypothesisTests, Printf

preq  = DataFrame(XLSX.readtable("data/pre-questionnaire.xlsx", 1))
postq = DataFrame(XLSX.readtable("data/post-questionnaire.xlsx", 1))

@assert size(preq)  == (19, 12)
@assert size(postq) == (18, 18)
[preq[!, i] .= convert.(Int, preq[:, i]) for i in 3:12]
[postq[!, i] .= convert.(Int, postq[:, i]) for i in 3:17]

sort!(preq, "ID:")
sort!(postq, "ID:")
preq1 = preq[Not(findall(==(false), [in(id, postq[:, "ID:"]) for id in preq[:, "ID:"]])), :]
@assert sum(preq1[:, "ID:"] .== postq[:, "ID:"]) == 18
n = size(preq1, 1)

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in [2, 4], j in 1:3]
    titles = "Q" .* string.(collect(1:6)) .* "\n "
    for i in 1:6
        p = @sprintf "%.2E" pvalue(SignedRankTest(preq1[:, i + 2], postq[:, i + 2]))
        if i <= 3
            scatter!(axs[1, mod1(i, 3)], vcat(zeros(n), ones(n)), vcat(preq1[:, i + 2], postq[:, i + 2]), color = ("black", 0.5))
            [lines!(axs[1, mod1(i, 3)], [0, 1], [preq1[j, i + 2], postq[j, i + 2]], color = ("#2774AE", 0.5)) for j in 1:n]
            text!(axs[1, mod1(i, 3)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), textsize = 15)
            Box(f[1, mod1(i, 3)], color = :gray90)
            Label(f[1, mod1(i, 3)], titles[i], tellwidth = false)
        else
            scatter!(axs[2, mod1(i, 3)], vcat(zeros(n), ones(n)), vcat(preq1[:, i + 2], postq[:, i + 2]), color = ("black", 0.5))
            [lines!(axs[2, mod1(i, 3)], [0, 1], [preq1[j, i + 2], postq[j, i + 2]], color = ("#2774AE", 0.5)) for j in 1:n]
            text!(axs[2, mod1(i, 3)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), textsize = 15)
            Box(f[3, mod1(i, 3)], color = :gray90)
            Label(f[3, mod1(i, 3)], titles[i], tellwidth = false)
        end
    end
    [xlims!(axs[i, j], -0.25, 1.25) for i in 1:2, j in 1:3]
    [axs[i, j].xticks = ([0, 1], ["pre", "post"]) for i in 1:2, j in 1:3]
    [hidexdecorations!(axs[i, j], ticklabels = false, grid = false) for i in 1:2, j in 1:3]
    Label(f[1:4, 0], text = "score", rotation = pi / 2)
    [rowgap!(f.layout, i, 2.5) for i in [1, 3]]
    save("figs/fig1a.pdf", f)
    f
end

begin
    f = Figure(resolution = (800, 800))
    axs = [Axis(f[i, j]) for i in [2, 4, 6], j in 1:2]
    titles = " " .* string.(collect(1:6)) .* "\n "
    for i in 1:6
        p = @sprintf "%.2E" pvalue(SignedRankTest(preq1[:, i + 2], postq[:, i + 2]))
        if i <= 2
            scatter!(axs[1, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(preq1[:, i + 2], postq[:, i + 2]), color = ("black", 0.5))
            [lines!(axs[1, mod1(i, 2)], [0, 1], [preq1[j, i + 2], postq[j, i + 2]], color = ("#2774AE", 0.5)) for j in 1:n]
            text!(axs[1, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), textsize = 15)
            Box(f[1, mod1(i, 2)], color = :gray90)
            Label(f[1, mod1(i, 2)], titles[i], tellwidth = false)
        elseif i <= 4
            scatter!(axs[2, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(preq1[:, i + 2], postq[:, i + 2]), color = ("black", 0.5))
            [lines!(axs[2, mod1(i, 2)], [0, 1], [preq1[j, i + 2], postq[j, i + 2]], color = ("#2774AE", 0.5)) for j in 1:n]
            text!(axs[2, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), textsize = 15)
            Box(f[3, mod1(i, 2)], color = :gray90)
            Label(f[3, mod1(i, 2)], titles[i], tellwidth = false)
        else
            scatter!(axs[3, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(preq1[:, i + 2], postq[:, i + 2]), color = ("black", 0.5))
            [lines!(axs[3, mod1(i, 2)], [0, 1], [preq1[j, i + 2], postq[j, i + 2]], color = ("#2774AE", 0.5)) for j in 1:n]
            text!(axs[3, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), textsize = 15)
            Box(f[5, mod1(i, 2)], color = :gray90)
            Label(f[5, mod1(i, 2)], titles[i], tellwidth = false)
        end
    end
    [xlims!(axs[i, j], -0.25, 1.25) for i in 1:3, j in 1:2]
    [axs[i, j].xticks = ([0, 1], ["pre", "post"]) for i in 1:3, j in 1:2]
    [hidexdecorations!(axs[i, j], ticklabels = false, grid = false) for i in 1:3, j in 1:2]
    Label(f[1:6, 0], text = "score", rotation = pi / 2)
    [rowgap!(f.layout, i, 2.5) for i in [1, 3, 5]]
    colsize!(f.layout, 1, Aspect(2, 1.5))
    colsize!(f.layout, 2, Aspect(2, 1.5))

    save("figs/fig1b.pdf", f)
    f
end

begin
    f = Figure()
    ax = Axis(f[1, 1])
    n1, n2 = count(==(3), preq[:, 9]), count(==(3), postq[:, 9])
    p1, p2 = n1 / 19, n2 / n
    barplot!(ax, [0, 1], [p1, p2], width = 0.5, color = ("royalblue", 0.8))
    errorbars!(ax, [0, 1], [p1, p2], sqrt.([p1 * (1 - p1) / n, p2 * (1 - p2) / n]))
    ax.xticks = ([0, 1], ["pre", "post"])
    ax.yticks = 0:0.2:0.95
    xlims!(ax, -0.5, 1.5)
    ylims!(ax, -0.05, 0.95)
    hidexdecorations!(ax, ticklabels = false, grid = false)
    p = @sprintf "%.2E" pvalue(FisherExactTest(n1, n2, 19 - n1, n - n2))
    text!(ax, 0.5, 0.925, text = "p = " * p, align = (:center, :top), textsize = 15)
    Label(f[1, 0], text = "proportion", rotation = pi / 2, tellheight = false)
    Box(f[0, 1], color = :gray90)
    Label(f[0, 1], "Q7", tellwidth = false)
    rowgap!(f.layout, 1, 2.5)
    save("figs/fig2.png", f)
    f
end

begin
    f = Figure()
    axs = [Axis(f[1, i]) for i in 1:3]
    titles = "Q" .* string.(collect(8:10)) .* " (pre)"
    # [(i, count(==(i), preq[!, 11])) for i in 1:5]
    for i in 1:3
        ind = findall(>(3), preq[:, 9 + i])
        n1 = count(==(3), preq[ind, 9])
        n2 = count(==(3), preq[Not(ind), 9])
        p1, p2 = n1 / length(ind), n2 / (19 - length(ind))
        barplot!(axs[i], [0, 1], [p2, p1], width = 0.5, color = ("royalblue", 0.8))
        errorbars!(axs[i], [0, 1], [p2, p1], sqrt.([p2 * (1 - p2) / (19 - length(ind)), p1 * (1 - p1) / length(ind)]))
        axs[i].xticks = ([0, 1], ["No", "Yes"])
        axs[i].yticks = 0:0.2:0.95
        xlims!(axs[i], -0.5, 1.5)
        ylims!(axs[i], -0.05, 0.95)
        p = @sprintf "%.2E" pvalue(FisherExactTest(n1, n2, length(ind) - n1, 19 - length(ind) - n2))
        text!(axs[i], 0.5, 0.925, text = "p = " * p, align = (:center, :top), textsize = 15)
        hidexdecorations!(axs[i], ticklabels = false, grid = false)
    end
    [Box(f[0, j], color = :gray90) for j in 1:3]
    [Label(f[0, j], titles[j], tellwidth = false) for j in 1:3]
    Label(f[1, 0], text = "proportion", rotation = pi / 2, tellheight = false)
    rowgap!(f.layout, 1, 2.5)
    save("figs/fig3.png", f)
    f
end

begin
    f = Figure(resolution = (800, 800))
    axs = [Axis(f[i, j]) for i in [2, 4, 6, 8], j in 1:2]
    titles = "Q" .* string.(collect(8:15)) .* " (post) \n "
    for i in 1:8
        row = div(i - 1, 2) + 1
        col = mod1(i, 2)
        ns = [count(==(s), postq[:, 9 + i]) for s in 1:5]
        barplot!(axs[row, col], 1:5, ns / n, color = ("#2774AE", 0.8))
        ylims!(axs[row, col], 0, 0.95)
        Box(f[2 * row - 1, col], color = :gray90)
        Label(f[2 * row - 1, col], titles[i], tellwidth = false)    
    end
    Label(f[end + 1, 1:2], text = "score")
    Label(f[1:8, 0], text = "proportion", rotation = pi / 2)
    [rowgap!(f.layout, i, 2.5) for i in [1, 3, 5, 7]]
    save("figs/fig4b.pdf", f)
    f
end

count(==("Hands on knot tying under the microscope"), skipmissing(postq[:, end]))