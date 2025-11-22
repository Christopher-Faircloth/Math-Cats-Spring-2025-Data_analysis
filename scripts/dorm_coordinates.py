import matplotlib.pyplot as plt
import matplotlib.image as mpimg

img = mpimg.imread("../data/images/bobcat village.png")
img = mpimg.imread("../data/images/main campus map.png")

fig, ax = plt.subplots(figsize=(10, 8))
ax.imshow(img)
ax.set_title("Click on the dorm locations")

coords = []

def onclick(event):
    if event.xdata and event.ydata:
        coords.append((event.xdata, event.ydata))
        ax.plot(event.xdata, event.ydata, 'ro')
        fig.canvas.draw()

cid = fig.canvas.mpl_connect('button_press_event', onclick)
plt.show()

print(coords)
# after clicking on the map i get these coordinates, i clicked the buildings left to right, top to bottom
#**************************not enough points************************** need to look at a better map, used a completely balank one
# (np.float64(875.0971612903224), np.float64(292.5833548387095)), 
# (np.float64(1032.4865806451612), np.float64(586.5033548387095)), 
# (np.float64(1062.8267096774193), np.float64(216.73303225806444)), 
# (np.float64(1117.818193548387), np.float64(578.918322580645)), 
# (np.float64(1193.6685161290322), np.float64(537.2006451612901)),
# (np.float64(1267.6225806451612), np.float64(192.08167741935472)), 
# (np.float64(1294.170193548387), np.float64(142.77896774193528)), 
# (np.float64(1366.2279999999998), np.float64(470.8316129032256)), 
# (np.float64(1423.1157419354838), np.float64(417.736387096774)), 
# (np.float64(1426.908258064516), np.float64(476.520387096774)), 
# (np.float64(1481.8997419354837), np.float64(639.5985806451612)), 
# (np.float64(1500.862322580645), np.float64(415.8401290322579)), 
# (np.float64(1519.8249032258063), np.float64(552.3707096774192)), 
# (np.float64(1586.1939354838707), np.float64(539.0969032258063))

# bobcat village
# [(np.float64(217.40016233766235), np.float64(212.48051948051943))]

# atempt 2
# [(np.float64(165.43979392551682), np.float64(244.4450063977372)), elena
# (np.float64(195.0498686780254), np.float64(313.5351808202572)), first five
# (np.float64(343.10024244056837), np.float64(565.2208162165803)), blanco
# (np.float64(397.385379486834), np.float64(570.1558286753317)), falls
# (np.float64(439.3329853862212), np.float64(636.778496868476)), sayers
# (np.float64(658.9410398006598), np.float64(557.8182975284532)), san marcos hall
# (np.float64(737.9012391406827), np.float64(493.6631355646846)),bexar hall
# (np.float64(713.2261768469255), np.float64(634.3109906391004)), richard a castro
# (np.float64(821.7964509394569), np.float64(244.4450063977372)), gaillardia hall
# (np.float64(851.4065256919655), np.float64(185.22485689272003)), chautauque hall
# (np.float64(876.0815879857228), np.float64(306.13266213213024)), college inn
# (np.float64(1026.5994679776413), np.float64(592.3633847397132)), department of hosing an residential life
# (np.float64(1058.6770489595256), np.float64(229.639969021483)), jackson hall
# (np.float64(1120.3647046939186), np.float64(599.7659034278404)), san jacinto hall
# (np.float64(1194.38989157519), np.float64(560.2858037578288)), tower hall
# (np.float64(1265.9475722270856), np.float64(209.8999191864773)), cibolo hall
# (np.float64(1295.5576469795942), np.float64(140.80974476395727)), alamito
# (np.float64(1379.4528587783686), np.float64(488.72812310593304)), retama
# (np.float64(1431.2704895952586), np.float64(493.6631355646846)), laurel
# (np.float64(1426.335477136507), np.float64(424.57296114216456)), mesquite
# (np.float64(1495.4256515590273), np.float64(412.2354299952858)), brogdon
# (np.float64(1485.5556266415242), np.float64(639.2460030978518)), sterry
# (np.float64(1522.56822008216), np.float64(560.2858037578288)), lantana
# (np.float64(1574.3858508990502), np.float64(547.9482726109503))] butler

#derrick hall
#[(np.float64(1281.1908387096776), np.float64(296.6184516129031))]