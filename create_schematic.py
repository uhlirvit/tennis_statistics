#!/usr/bin/env python3
"""
Generate a schematic figure for CFD simulation of flow over a wall-mounted heated cube.
Based on the experimental setup of Meinders et al. and numerical setup dimensions.

Domain dimensions:
- Hc (cube height): 15 mm
- X length: 315 mm (21Hc)
- Y width: 165 mm (11Hc)
- Z height: 50 mm (3.33Hc)

The cube is placed at a distance from the inlet (approximately 5Hc upstream region).
"""

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyArrowPatch, Rectangle, FancyBboxPatch, Polygon
import numpy as np

# Set up the figure with two subplots
fig = plt.figure(figsize=(12, 5))

# Use a clean, publication-ready style
plt.rcParams.update({
    'font.family': 'serif',
    'font.size': 10,
    'axes.linewidth': 0.8,
    'lines.linewidth': 0.8,
    'mathtext.fontset': 'dejavuserif',
})

# ============================================================================
# Subplot (a): 3D perspective view - showing front and side walls instead
# ============================================================================
ax1 = fig.add_subplot(121)
ax1.set_aspect('equal')
ax1.axis('off')

# Perspective projection parameters - showing domain from front-right
angle = 30  # degrees
cos_a = np.cos(np.radians(angle))
sin_a = np.sin(np.radians(angle))

# Domain dimensions in Hc units (Hc = cube height = 15 mm)
L_x = 21  # length in Hc
L_y = 11  # width in Hc  
L_z = 3.33  # height in Hc
Hc = 1  # normalized cube height

# Cube position
cube_x = 5  # x position of cube front face

# Scale factor for drawing
scale = 0.12

def iso_transform(x, y, z):
    """Transform 3D coordinates to 2D isometric-like projection.
    View from front-right: we see the front (y=0) and right (x=Lx) faces."""
    x_2d = (x + y * cos_a) * scale
    y_2d = (z + y * sin_a) * scale
    return x_2d, y_2d

# Draw the bottom plane (floor)
floor_corners = [
    iso_transform(0, 0, 0),
    iso_transform(L_x, 0, 0),
    iso_transform(L_x, L_y, 0),
    iso_transform(0, L_y, 0),
]
floor = patches.Polygon(floor_corners, fill=True, facecolor='#f0f0f0', 
                         edgecolor='black', linewidth=0.8, zorder=1)
ax1.add_patch(floor)

# Draw the back wall (y = L_y) - this is visible from front
back_wall_corners = [
    iso_transform(0, L_y, 0),
    iso_transform(L_x, L_y, 0),
    iso_transform(L_x, L_y, L_z),
    iso_transform(0, L_y, L_z),
]
back_wall = patches.Polygon(back_wall_corners, fill=True, facecolor='#e0e0e0',
                             edgecolor='black', linewidth=0.8, zorder=2)
ax1.add_patch(back_wall)

# Draw the right wall (outlet, x = L_x) - partially visible
right_wall_corners = [
    iso_transform(L_x, 0, 0),
    iso_transform(L_x, L_y, 0),
    iso_transform(L_x, L_y, L_z),
    iso_transform(L_x, 0, L_z),
]
right_wall = patches.Polygon(right_wall_corners, fill=True, facecolor='#d0d0d0',
                              edgecolor='black', linewidth=0.8, zorder=2)
ax1.add_patch(right_wall)

# Draw front edge of floor (y=0 line)
ax1.plot([iso_transform(0, 0, 0)[0], iso_transform(L_x, 0, 0)[0]],
         [iso_transform(0, 0, 0)[1], iso_transform(L_x, 0, 0)[1]], 
         'k-', linewidth=1.0, zorder=5)

# Draw the cube (positioned at center of width)
cube_y = L_y/2 - 0.5  # y position of cube front face

# Cube front face (y = cube_y)
cube_front = [
    iso_transform(cube_x, cube_y, 0),
    iso_transform(cube_x + Hc, cube_y, 0),
    iso_transform(cube_x + Hc, cube_y, Hc),
    iso_transform(cube_x, cube_y, Hc),
]
cube_front_patch = patches.Polygon(cube_front, fill=True, facecolor='#cc4444',
                                    edgecolor='black', linewidth=1.0, zorder=10)
ax1.add_patch(cube_front_patch)

# Cube right face (x = cube_x + Hc)
cube_right = [
    iso_transform(cube_x + Hc, cube_y, 0),
    iso_transform(cube_x + Hc, cube_y + Hc, 0),
    iso_transform(cube_x + Hc, cube_y + Hc, Hc),
    iso_transform(cube_x + Hc, cube_y, Hc),
]
cube_right_patch = patches.Polygon(cube_right, fill=True, facecolor='#aa3333',
                                    edgecolor='black', linewidth=1.0, zorder=10)
ax1.add_patch(cube_right_patch)

# Cube top face
cube_top = [
    iso_transform(cube_x, cube_y, Hc),
    iso_transform(cube_x + Hc, cube_y, Hc),
    iso_transform(cube_x + Hc, cube_y + Hc, Hc),
    iso_transform(cube_x, cube_y + Hc, Hc),
]
cube_top_patch = patches.Polygon(cube_top, fill=True, facecolor='#dd5555',
                                  edgecolor='black', linewidth=1.0, zorder=10)
ax1.add_patch(cube_top_patch)

# Add flow direction arrow - positioned clearly in front
arrow_start = iso_transform(-1, L_y/2, L_z/2)
arrow_end = iso_transform(2, L_y/2, L_z/2)
ax1.annotate('', xy=arrow_end, xytext=arrow_start,
             arrowprops=dict(arrowstyle='->', color='blue', lw=1.5),
             zorder=20)
ax1.text(arrow_start[0] - 0.08, arrow_start[1], 'Flow', fontsize=10, 
         ha='right', va='center', color='blue')

# Add dimension annotations
# Length (21Hc) - along bottom front edge
p1 = iso_transform(0, 0, 0)
p2 = iso_transform(L_x, 0, 0)
dim_offset = -0.18
ax1.annotate('', xy=(p2[0], p2[1] + dim_offset), xytext=(p1[0], p1[1] + dim_offset),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax1.text((p1[0] + p2[0])/2, (p1[1] + p2[1])/2 + dim_offset - 0.1, r'21$H_\mathrm{c}$ (315 mm)',
         fontsize=9, ha='center', va='top')

# Width (11Hc) - along the depth direction at outlet
p1 = iso_transform(L_x, 0, 0)
p2 = iso_transform(L_x, L_y, 0)
ax1.annotate('', xy=(p2[0] + 0.08, p2[1] + 0.05), xytext=(p1[0] + 0.08, p1[1] + 0.05),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
mid_x = (p1[0] + p2[0])/2
mid_y = (p1[1] + p2[1])/2
ax1.text(mid_x + 0.22, mid_y + 0.08, r'11$H_\mathrm{c}$' + '\n(165 mm)', fontsize=9, ha='left', va='center')

# Height (3.3Hc) - vertical at back-right corner
p1 = iso_transform(L_x, L_y, 0)
p2 = iso_transform(L_x, L_y, L_z)
ax1.annotate('', xy=(p2[0] + 0.08, p2[1]), xytext=(p1[0] + 0.08, p1[1]),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax1.text(p2[0] + 0.15, (p1[1] + p2[1])/2, r'3.3$H_\mathrm{c}$' + '\n(50 mm)', fontsize=9, 
         ha='left', va='center')

# Cube height annotation - on the front-left of the cube
cube_dim_x = cube_x - 0.3
cube_dim_y = cube_y
p1 = iso_transform(cube_dim_x, cube_dim_y, 0)
p2 = iso_transform(cube_dim_x, cube_dim_y, Hc)
ax1.annotate('', xy=(p2[0], p2[1]), xytext=(p1[0], p1[1]),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax1.text(p1[0] - 0.06, (p1[1] + p2[1])/2, r'$H_\mathrm{c}$', fontsize=9, 
         ha='right', va='center')

# Label for cube - positioned below and in front, not crossing any boundaries
label_pos = iso_transform(cube_x + 0.5, cube_y - 2.5, 0)
ax1.text(label_pos[0], label_pos[1] - 0.05, 'Cube', fontsize=9, ha='center', va='top')
# Arrow pointing to cube
cube_center = iso_transform(cube_x + 0.5, cube_y + 0.5, Hc/2)
ax1.annotate('', xy=cube_center,
             xytext=(label_pos[0], label_pos[1]),
             arrowprops=dict(arrowstyle='->', color='black', lw=0.6))

# Coordinate system - positioned in bottom-left front
origin = iso_transform(-0.5, -1, 0)
axis_len = 0.22
ax1.annotate('', xy=(origin[0] + axis_len, origin[1]), xytext=origin,
             arrowprops=dict(arrowstyle='->', color='black', lw=0.8))
ax1.annotate('', xy=(origin[0] + axis_len*cos_a, origin[1] + axis_len*sin_a), xytext=origin,
             arrowprops=dict(arrowstyle='->', color='black', lw=0.8))
ax1.annotate('', xy=(origin[0], origin[1] + axis_len), xytext=origin,
             arrowprops=dict(arrowstyle='->', color='black', lw=0.8))
ax1.text(origin[0] + axis_len + 0.04, origin[1], 'x', fontsize=9, ha='left', va='center')
ax1.text(origin[0] + axis_len*cos_a + 0.04, origin[1] + axis_len*sin_a, 'y', fontsize=9, ha='left', va='bottom')
ax1.text(origin[0], origin[1] + axis_len + 0.04, 'z', fontsize=9, ha='center', va='bottom')

ax1.set_xlim(-0.5, 3.2)
ax1.set_ylim(-0.6, 1.4)
ax1.set_title('(a)', fontsize=11, pad=10, loc='left', fontweight='bold')

# ============================================================================
# Subplot (b): Side view (x-z plane) - made taller/larger
# ============================================================================
ax2 = fig.add_subplot(122)
ax2.set_aspect('equal')
ax2.axis('off')

# Scale for side view - increased vertical emphasis
s_x = 0.10  # horizontal scale
s_z = 0.18  # vertical scale (stretched for clarity)

# Channel walls
# Bottom wall
ax2.plot([0, L_x * s_x], [0, 0], 'k-', linewidth=1.5)
# Top wall  
ax2.plot([0, L_x * s_x], [L_z * s_z, L_z * s_z], 'k-', linewidth=1.5)

# Inlet
ax2.plot([0, 0], [0, L_z * s_z], 'k-', linewidth=1.0)
# Outlet
ax2.plot([L_x * s_x, L_x * s_x], [0, L_z * s_z], 'k-', linewidth=1.0)

# Draw the cube
cube_rect = patches.Rectangle((cube_x * s_x, 0), Hc * s_x, Hc * s_z,
                                facecolor='#cc4444', edgecolor='black', 
                                linewidth=1.0, zorder=5)
ax2.add_patch(cube_rect)

# Hatching for walls (to indicate solid boundaries)
hatch_len = 0.04
hatch_spacing = 0.08
for i in range(int(L_x * s_x / hatch_spacing) + 2):
    x_pos = i * hatch_spacing
    if x_pos <= L_x * s_x:
        # Bottom wall hatching
        ax2.plot([x_pos, x_pos + hatch_len], [-hatch_len, 0], 'k-', linewidth=0.5)
        # Top wall hatching
        ax2.plot([x_pos, x_pos + hatch_len], [L_z * s_z, L_z * s_z + hatch_len], 'k-', linewidth=0.5)

# Realistic velocity profile at inlet - developed turbulent profile with boundary layers
# Based on user's image: thin laminar BL at top (H3), uniform core (H2), thicker turbulent BL at bottom (H1)
n_arrows = 20
z_positions = np.linspace(0.02 * L_z * s_z, 0.98 * L_z * s_z, n_arrows)

def velocity_profile(z_norm):
    """
    Create a realistic developed turbulent channel flow profile.
    z_norm: normalized position (0 = bottom, 1 = top)
    
    Structure:
    - Bottom: thicker turbulent boundary layer (H1)
    - Middle: uniform core (H2)
    - Top: thin laminar sublayer (H3)
    """
    # Approximate 1/7 power law for turbulent BL, with asymmetry
    # Bottom BL is thicker (turbulent), top BL is thinner
    
    if z_norm < 0.25:  # Bottom turbulent BL (H1) - thicker
        # Power law growth
        u = (z_norm / 0.25) ** (1/7)
    elif z_norm > 0.85:  # Top laminar BL (H3) - thinner
        # Steeper drop near wall
        u = ((1 - z_norm) / 0.15) ** 0.5
    else:  # Uniform core (H2)
        u = 1.0
    
    return u * 0.25  # Scale for arrow length

# Draw velocity profile with arrows
inlet_x = -0.08
profile_x = []
profile_z = []

for z in z_positions:
    z_norm = z / (L_z * s_z)
    u = velocity_profile(z_norm)
    if u > 0.02:  # Only draw if velocity is significant
        ax2.arrow(inlet_x, z, u, 0, head_width=0.018, head_length=0.015, 
                  fc='blue', ec='blue', linewidth=0.5, zorder=15)
    profile_x.append(inlet_x + u)
    profile_z.append(z)

# Draw the profile curve
ax2.plot(profile_x, profile_z, 'b-', linewidth=1.0, zorder=14)

# Add velocity profile labels
# H1 - turbulent boundary layer (bottom)
h1_z = 0.12 * L_z * s_z
ax2.annotate('', xy=(inlet_x - 0.02, 0), xytext=(inlet_x - 0.02, 0.25 * L_z * s_z),
             arrowprops=dict(arrowstyle='<->', color='gray', lw=0.6))
ax2.text(inlet_x - 0.06, 0.125 * L_z * s_z, r'$H_1$', fontsize=8, ha='right', va='center', color='gray')

# H2 - uniform core
ax2.annotate('', xy=(inlet_x - 0.02, 0.25 * L_z * s_z), xytext=(inlet_x - 0.02, 0.85 * L_z * s_z),
             arrowprops=dict(arrowstyle='<->', color='gray', lw=0.6))
ax2.text(inlet_x - 0.06, 0.55 * L_z * s_z, r'$H_2$', fontsize=8, ha='right', va='center', color='gray')

# H3 - laminar boundary layer (top)
ax2.annotate('', xy=(inlet_x - 0.02, 0.85 * L_z * s_z), xytext=(inlet_x - 0.02, L_z * s_z),
             arrowprops=dict(arrowstyle='<->', color='gray', lw=0.6))
ax2.text(inlet_x - 0.06, 0.925 * L_z * s_z, r'$H_3$', fontsize=8, ha='right', va='center', color='gray')

# Dimension annotations
# Total length
ax2.annotate('', xy=(L_x * s_x, -0.18), xytext=(0, -0.18),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax2.text(L_x/2 * s_x, -0.26, r'21$H_\mathrm{c}$ (315 mm)', fontsize=9, ha='center', va='top')

# Upstream distance to cube (5Hc)
ax2.annotate('', xy=(cube_x * s_x, -0.08), xytext=(0, -0.08),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax2.text(cube_x/2 * s_x, -0.13, r'5$H_\mathrm{c}$', fontsize=8, ha='center', va='top')

# Downstream distance (15Hc)
ax2.annotate('', xy=(L_x * s_x, -0.08), xytext=((cube_x + Hc) * s_x, -0.08),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax2.text((cube_x + Hc + (L_x - cube_x - Hc)/2) * s_x, -0.13, r'15$H_\mathrm{c}$', fontsize=8, 
         ha='center', va='top')

# Channel height
ax2.annotate('', xy=(L_x * s_x + 0.12, L_z * s_z), xytext=(L_x * s_x + 0.12, 0),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax2.text(L_x * s_x + 0.18, L_z/2 * s_z, r'3.3$H_\mathrm{c}$' + '\n(50 mm)', fontsize=9, 
         ha='left', va='center')

# Cube height  
ax2.annotate('', xy=((cube_x + Hc) * s_x + 0.06, Hc * s_z), 
             xytext=((cube_x + Hc) * s_x + 0.06, 0),
             arrowprops=dict(arrowstyle='<->', color='black', lw=0.8))
ax2.text((cube_x + Hc) * s_x + 0.10, Hc/2 * s_z, r'$H_\mathrm{c}$' + '\n(15 mm)', fontsize=9, 
         ha='left', va='center')

# Cube label - positioned above, not crossing boundaries
ax2.text(cube_x * s_x + Hc/2 * s_x, Hc * s_z + 0.1, 'Cube', fontsize=9, 
         ha='center', va='bottom', color='black')

# Inlet/Outlet labels
ax2.text(0, L_z * s_z + 0.08, 'Inlet', fontsize=9, ha='center', va='bottom')
ax2.text(L_x * s_x, L_z * s_z + 0.08, 'Outlet', fontsize=9, ha='center', va='bottom')

ax2.set_xlim(-0.35, L_x * s_x + 0.4)
ax2.set_ylim(-0.4, L_z * s_z + 0.2)
ax2.set_title('(b)', fontsize=11, pad=10, loc='left', fontweight='bold')

# Adjust layout and save
plt.tight_layout()

# Save as PDF (vector format, ideal for LaTeX)
plt.savefig('/home/claude/domain_schematic.pdf', format='pdf', 
            bbox_inches='tight', dpi=300)

# Also save as PNG for preview
plt.savefig('/home/claude/domain_schematic.png', format='png', 
            bbox_inches='tight', dpi=300)

print("Figures saved:")
print("  - domain_schematic.pdf (for LaTeX)")
print("  - domain_schematic.png (for preview)")
