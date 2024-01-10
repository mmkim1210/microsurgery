using XLSX, DataFrames, CairoMakie, HypothesisTests, Printf, Statistics

score  = DataFrame(XLSX.readtable("data/score.xlsx", 2))
meta = DataFrame(XLSX.readtable("data/score.xlsx", 4))
[score[!, i] .= convert.(Int, score[:, i]) for i in [1, 3:10]]
[meta[!, i] .= convert.(Int, meta[:, i]) for i in [1, 3]]
n = size(meta, 1)

df = select(score, Not(2))
df = combine(groupby(df, 1), names(df)[Not(1)] .=> mean, renamecols = false)

pre = df[findall(isequal(meta[1, begin]), df[:, begin]), :]
post = df[findall(isequal(meta[1, 3]), df[:, begin]), :]
for i in 2:n
    pre = vcat(pre, df[findall(isequal(meta[i, begin]), df[:, begin]), :])
    post = vcat(post, df[findall(isequal(meta[i, 3]), df[:, begin]), :])
end

meta
score

CairoMakie.activate!(type = "svg")

begin
    f = Figure()
    axs = [Axis(f[i, j]) for i in [2, 4], j in 1:3]
    titles = ["Respect for tissue", "Instrument handling", "Knot tying", "Grabbing edge of vessel", "Grabbing needle tip", 
        "Hand crossing during tying"]
    for i in 1:6
        p = @sprintf "%.2E" pvalue(SignedRankTest(pre[:, i + 1], post[:, i + 1]))
        if i <= 3
            scatter!(axs[1, mod1(i, 3)], vcat(zeros(n), ones(n)), vcat(pre[:, i + 1], post[:, i + 1]), color = ("black", 0.5))
            [lines!(axs[1, mod1(i, 3)], [0, 1], [pre[j, i + 1], post[j, i + 1]], color = ("black", 0.25)) for j in 1:n] # #CB3C33
            text!(axs[1, mod1(i, 3)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
            Box(f[1, mod1(i, 3)], color = :gray90)
            Label(f[1, mod1(i, 3)], titles[i], tellwidth = false)
        else
            scatter!(axs[2, mod1(i, 3)], vcat(zeros(n), ones(n)), vcat(pre[:, i + 1], post[:, i + 1]), color = ("black", 0.5))
            [lines!(axs[2, mod1(i, 3)], [0, 1], [pre[j, i + 1], post[j, i + 1]], color = ("black", 0.25)) for j in 1:n] # #CB3C33
            text!(axs[2, mod1(i, 3)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
            Box(f[3, mod1(i, 3)], color = :gray90)
            Label(f[3, mod1(i, 3)], titles[i], tellwidth = false)
        end
    end
    [ylims!(axs[i, j], 0.75, 5.25) for i in 1:2, j in 1:3]
    [xlims!(axs[i, j], -0.25, 1.25) for i in 1:2, j in 1:3]
    [axs[i, j].xticks = ([0, 1], ["pre", "post"]) for i in 1:2, j in 1:3]
    [hidexdecorations!(axs[i, j], ticklabels = false) for i in 1:2, j in 1:3]
    [hideydecorations!(axs[i, j], ticklabels = false, ticks = false) for i in 1:2, j in 1:3]

    ax = Axis(f[6, 1])
    p = @sprintf "%.2E" pvalue(SignedRankTest(pre[:, 8], post[:, 8]))
    scatter!(ax, vcat(zeros(n), ones(n)), vcat(pre[:, 8], post[:, 8]), color = ("black", 0.5))
    [lines!(ax, [0, 1], [pre[j, 8], post[j, 8]], color = ("black", 0.25)) for j in 1:n]
    text!(ax, 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
    ylims!(ax, 0.75, 5.25)
    xlims!(ax, -0.25, 1.25)
    ax.xticks = ([0, 1], ["pre", "post"])
    hidexdecorations!(ax, ticklabels = false)
    hideydecorations!(ax, ticklabels = false, ticks = false)

    Box(f[5, 1], color = :gray90)
    Label(f[5, 1], "Suture breaking", tellwidth = false)

    Label(f[1:6, 0], text = "score", rotation = pi / 2)
    [rowgap!(f.layout, i, 2.5) for i in [1, 3, 5]]
    save("figs/fig5a.svg", f)
    f
end

begin
    f = Figure(resolution = (800, 500))
    axs = [Axis(f[i, j]) for i in [2, 4, 6], j in 1:2]
    titles = ["Respect for tissue", "Instrument handling", "Knot tying", "Grabbing edge of vessel", "Grabbing needle tip", 
        "Hand crossing during tying"]
    for i in 1:6
        p = @sprintf "%.2E" pvalue(SignedRankTest(pre[:, i + 1], post[:, i + 1]))
        if i <= 2
            scatter!(axs[1, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(pre[:, i + 1], post[:, i + 1]), color = ("black", 0.5))
            [lines!(axs[1, mod1(i, 2)], [0, 1], [pre[j, i + 1], post[j, i + 1]], color = ("black", 0.25)) for j in 1:n] # #CB3C33
            text!(axs[1, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
            Box(f[1, mod1(i, 2)], color = :gray90)
            Label(f[1, mod1(i, 2)], titles[i], tellwidth = false)
        elseif i <= 4
            scatter!(axs[2, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(pre[:, i + 1], post[:, i + 1]), color = ("black", 0.5))
            [lines!(axs[2, mod1(i, 2)], [0, 1], [pre[j, i + 1], post[j, i + 1]], color = ("black", 0.25)) for j in 1:n] # #CB3C33
            text!(axs[2, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
            Box(f[3, mod1(i, 2)], color = :gray90)
            Label(f[3, mod1(i, 2)], titles[i], tellwidth = false)
        else
            scatter!(axs[3, mod1(i, 2)], vcat(zeros(n), ones(n)), vcat(pre[:, i + 1], post[:, i + 1]), color = ("black", 0.5))
            [lines!(axs[3, mod1(i, 2)], [0, 1], [pre[j, i + 1], post[j, i + 1]], color = ("black", 0.25)) for j in 1:n] # #CB3C33
            text!(axs[3, mod1(i, 2)], 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
            Box(f[5, mod1(i, 2)], color = :gray90)
            Label(f[5, mod1(i, 2)], titles[i], tellwidth = false)
        end
    end
    [ylims!(axs[i, j], 0.75, 5.25) for i in 1:3, j in 1:2]
    [xlims!(axs[i, j], -0.25, 1.25) for i in 1:3, j in 1:2]
    [axs[i, j].xticks = ([0, 1], ["pre", "post"]) for i in 1:3, j in 1:2]
    [hidexdecorations!(axs[i, j], ticklabels = false, grid = false) for i in 1:3, j in 1:2]

    ax = Axis(f[8, 1])
    p = @sprintf "%.2E" pvalue(SignedRankTest(pre[:, 8], post[:, 8]))
    scatter!(ax, vcat(zeros(n), ones(n)), vcat(pre[:, 8], post[:, 8]), color = ("black", 0.25))
    [lines!(ax, [0, 1], [pre[j, 8], post[j, 8]], color = ("black", 0.25)) for j in 1:n]
    text!(ax, 0.85, 1.5, text = "p = " * p, align = (:center, :top), fontsize = 10)
    ylims!(ax, 0.75, 5.25)
    xlims!(ax, -0.25, 1.25)
    ax.xticks = ([0, 1], ["pre", "post"])
    hidexdecorations!(ax, ticklabels = false, grid = false)

    Box(f[7, 1], color = :gray90)
    Label(f[7, 1], "Suture breaking", tellwidth = false)

    Label(f[1:8, 0], text = "score", rotation = pi / 2)
    [rowgap!(f.layout, i, 2.5) for i in [1, 3, 5, 7]]
    colsize!(f.layout, 1, Aspect(2, 2))
    colsize!(f.layout, 2, Aspect(2, 2))

    save("figs/fig5b.svg", f)
    f
end