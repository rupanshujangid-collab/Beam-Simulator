# Beam calculations
beam_length = 2.0      # meters
load = 500             # Newtons
E = 200e9              # Young's Modulus - Steel (Pa)
I = 8.33e-6            # Moment of Inertia (m^4)

# Reactions (simply supported, central load)
RA = load / 2
RB = load / 2

# Max Bending Moment
M_max = (load * beam_length) / 4

# Max Deflection
y_max = (load * beam_length**3) / (48 * E * I)

print("Reaction at A:", RA, "N")
print("Reaction at B:", RB, "N")
print("Max Bending Moment:", M_max, "N.m")
print("Max Deflection:", round(y_max * 1000, 4), "mm")