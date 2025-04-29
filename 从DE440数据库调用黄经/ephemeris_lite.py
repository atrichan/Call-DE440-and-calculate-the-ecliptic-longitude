import sys
from datetime import timezone, timedelta
import pandas as pd
from skyfield.api import load
from skyfield import almanac

# ---------- 常量 ----------
CST = timezone(timedelta(hours=8), name='CST')
ts = load.timescale(builtin=False)   # ← 关闭内置表，自动用 IERS 实测 ΔT


# 首次运行会自动下载 DE440 星历文件
eph = load('de440.bsp')
sun, moon, earth = eph['sun'], eph['moon'], eph['earth']

# ---------- 工具 ----------
def ecl_lon(body, t):
    """真黄经 (0–360°)，仅含章动；无光行差/视差"""
    # 纯几何向量 (body-earth)，不含光行时差
    vec = (body - earth).at(t)
    _, lon, _ = vec.ecliptic_latlon(epoch='date')
    return lon.degrees % 360

def ang_diff(a, b):
    """最短角差 (-180°, +180°]"""
    return ((a - b + 180) % 360) - 180

def find_events(fn_sign, fn_val, target, t0, t1,
                step_h=1, tol=0.01):
    """返回所有满足 |fn_val(t)-target|<tol 的时刻列表"""
    fn_sign.step_days = step_h / 24
    times, _ = almanac.find_discrete(t0, t1, fn_sign)

    # 二次筛选，排除 ±180° 翻号造成的假根
    return [t for t in times
            if abs(ang_diff(fn_val(t), target)) < tol]

def dms(deg):
    d = int(deg)
    m = int((deg - d) * 60)
    s = (deg - d - m / 60) * 3600
    sign = '-' if deg < 0 else ''
    return f"{sign}{abs(d):3d}°{abs(m):02d}′{abs(s):04.1f}″"

# ---------- 主程 ----------
def main():
    # --- 输入 ---
    start_in = pd.to_datetime(input("开始时间（北京时间，如 2025-04-01 00:00:00）："))
    end_in   = pd.to_datetime(input("结束时间（北京时间，如 2025-05-01 00:00:00）："))

    if start_in.tzinfo is None: start_in = start_in.tz_localize(CST)
    if end_in.tzinfo   is None: end_in   = end_in.tz_localize(CST)
    if start_in >= end_in:
        sys.exit("结束时间必须晚于开始时间！")

    # 转为 Skyfield 时间
    t0, t1 = ts.from_datetime(start_in), ts.from_datetime(end_in)

    # --- 初始量 ---
    λs0 = ecl_lon(sun,  t0)
    λm0 = ecl_lon(moon, t0)
    Δ0  = (λm0 - λs0) % 360

    # --- 构造符号函数 ---
    f_s  = lambda t: ang_diff(ecl_lon(sun,  t), λs0)  > 0
    f_m  = lambda t: ang_diff(ecl_lon(moon, t), λm0)  > 0
    f_dm = lambda t: ang_diff((ecl_lon(moon, t) - ecl_lon(sun, t)) % 360, Δ0) > 0

    print("\n正在计算，请稍候…")

    ev_s  = find_events(f_s,  lambda t: ecl_lon(sun,  t),             λs0, t0, t1)
    ev_m  = find_events(f_m,  lambda t: ecl_lon(moon, t),             λm0, t0, t1)
    ev_dm = find_events(f_dm, lambda t: (ecl_lon(moon, t) - ecl_lon(sun,  t)) % 360, Δ0,  t0, t1)

    # --- 收集数据 ---
    data = []

    def collect_data(title, arr):
        data.append([f"{title}（共 {len(arr)} 次）", "", "", ""])  # 添加标题行
        for t in arr:
            beijing_dt = t.utc_datetime().astimezone(CST)
            λs = ecl_lon(sun,  t)
            λm = ecl_lon(moon, t)
            Δ  = (λm - λs) % 360
            data.append([beijing_dt.strftime("%Y-%m-%d %H:%M:%S"), dms(λs), dms(λm), dms(Δ)])

    collect_data("► 太阳黄经回到首日值的时刻",  ev_s)
    collect_data("► 月亮黄经回到首日值的时刻", ev_m)
    collect_data("► 日月黄经差回到首日值的时刻", ev_dm)

    # --- 保存数据到 Excel ---
    filename = input('请输入导出的excel的名称：') + '.xlsx'
    df = pd.DataFrame(data, columns=["时间", "太阳黄经", "月亮黄经", "黄经差"])
    df.to_excel(filename, index=False)

    print(f"\n数据已保存到 {filename}")
    sys.stdout.write('\n')
    input('加载完毕，请按Enter键退出')

if __name__ == "__main__":
    main()

